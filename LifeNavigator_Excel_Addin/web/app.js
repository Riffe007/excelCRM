/* global Office, Excel, d3 */
const SHEETS = { leads: "Leads", activities: "Activities", accounts: "Accounts" };
const TABLES = { leads: "LeadsTable", activities: "ActivitiesTable", accounts: "AccountsTable" };

Office.onReady(() => {
  initUI();
  document.getElementById("refreshBtn").addEventListener("click", refreshAll);
  document.getElementById("leadForm").addEventListener("submit", onSubmit);
  document.getElementById("clearBtn").addEventListener("click", () => document.getElementById("leadForm").reset());
  document.getElementById("searchBtn").addEventListener("click", onSearch);
  document.getElementById("showAllBtn").addEventListener("click", refreshLeadsTable);
  document.getElementById("logActivityBtn").addEventListener("click", logActivityPrompt);
  refreshAll();
});

function initUI(){
  document.querySelectorAll(".tab").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach(b=>b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach(p=>p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(btn.dataset.tab).classList.add("active");
    });
  });
}

async function refreshAll(){
  await ensureTables();
  const data = await getAllData();
  renderKPIs(data);
  renderStatusChart(data.leads);
  renderStageValueChart(data.leads);
  renderMonthChart(data.leads);
  populateLeadsTable(data.leads);
  populateActivitiesTable(data.activities);
}

// Ensure sheets & tables exist
async function ensureTables(){
  await Excel.run(async (ctx) => {
    const wb = ctx.workbook;

    function ensureSheet(name){
      const sheet = wb.worksheets.getItemOrNullObject(name);
      sheet.load("name");
      return sheet;
    }
    const leads = ensureSheet(SHEETS.leads);
    const acts  = ensureSheet(SHEETS.activities);
    const accts = ensureSheet(SHEETS.accounts);
    await ctx.sync();

    const created = [];
    if (leads.isNullObject) {
      const s = wb.worksheets.add(SHEETS.leads);
      s.getRange("A1:Q1").values = [[
        "ID","Created On","Last Updated","Owner","Account","Name","Title/Role","Email","Phone","City/State","Source","Priority","Status","Stage","Est. Value ($)","Close Date","Notes"
      ]];
      created.push(s);
    }
    if (acts.isNullObject) {
      const s = wb.worksheets.add(SHEETS.activities);
      s.getRange("A1:G1").values = [["Timestamp","Lead ID","Owner","Type","Notes","Next Step","Due Date"]];
      created.push(s);
    }
    if (accts.isNullObject) {
      const s = wb.worksheets.add(SHEETS.accounts);
      s.getRange("A1:I1").values = [["Account ID","Account Name","Type","Owner","City/State","Website","Priority","Status","Notes"]];
      created.push(s);
    }
    await ctx.sync();

    // Add tables if missing
    function ensureTable(sheetName, tableName, address){
      const sheet = wb.worksheets.getItem(sheetName);
      const tables = wb.tables;
      let table = tables.getItemOrNullObject(tableName);
      table.load("name");
      return ctx.sync().then(()=>{
        if (table.isNullObject) {
          table = tables.add(sheet.getRange(address), true /*hasHeaders*/);
          table.name = tableName;
        }
        return ctx.sync();
      });
    }

    await ensureTable(SHEETS.leads, TABLES.leads, "A1:Q1");
    await ensureTable(SHEETS.activities, TABLES.activities, "A1:G1");
    await ensureTable(SHEETS.accounts, TABLES.accounts, "A1:I1");

    created.forEach(s => s.activate());
    if (created.length) await ctx.sync();
  });
}

// Load data from tables
async function getAllData(){
  return Excel.run(async (ctx) => {
    const leadsTable = ctx.workbook.tables.getItem(TABLES.leads);
    const actsTable  = ctx.workbook.tables.getItem(TABLES.activities);

    const leadsRange = leadsTable.getDataBodyRange().load(["values", "rowCount", "columnCount"]);
    const actsRange  = actsTable.getDataBodyRange().load(["values"]);
    await ctx.sync();

    const leadCols = ["ID","Created On","Last Updated","Owner","Account","Name","Title/Role","Email","Phone","City/State","Source","Priority","Status","Stage","Est. Value ($)","Close Date","Notes"];
    const leads = (leadsRange.values || []).map(row => Object.fromEntries(row.map((v,i)=>[leadCols[i], v])));
    const activities = (actsRange.values || []).map(row => ({
      "Timestamp": row[0], "Lead ID": row[1], "Owner": row[2], "Type": row[3], "Notes": row[4], "Next Step": row[5], "Due Date": row[6]
    }));
    return { leads, activities };
  });
}

// Submit form -> append to Leads table
async function onSubmit(e){
  e.preventDefault();
  const fd = new FormData(e.target);
  const obj = Object.fromEntries(fd.entries());
  obj.value = obj.value ? parseFloat(obj.value) : "";
  await Excel.run(async (ctx) => {
    const leadsTable = ctx.workbook.tables.getItem(TABLES.leads);
    // compute next ID
    let nextId = 1;
    try {
      const idRange = leadsTable.columns.getItem("ID").getDataBodyRange().load("values");
      await ctx.sync();
      const ids = (idRange.values||[]).flat().filter(v=>v!=="" && !isNaN(v));
      if (ids.length) nextId = Math.max(...ids.map(Number)) + 1;
    } catch {}

    const now = new Date().toISOString().slice(0,10);
    const row = [
      nextId, now, now, obj.owner||"", obj.account||"", obj.name||"", obj.title||"", obj.email||"", obj.phone||"",
      obj.citystate||"", obj.source||"", obj.priority||"", obj.status||"New", obj.stage||"Discovery",
      obj.value||"", obj.closedate||"", obj.notes||""
    ];
    leadsTable.rows.add(null, [row]);
    await ctx.sync();
  });
  e.target.reset();
  await refreshAll();
}

// Search in table
async function onSearch(){
  const q = (document.getElementById("searchInput").value||"").toLowerCase();
  if (!q) { refreshLeadsTable(); return; }
  const data = await getAllData();
  const filtered = data.leads.filter(r => (r["Name"]||"").toLowerCase().includes(q) || (r["Email"]||"").toLowerCase().includes(q));
  populateLeadsTable(filtered);
}

// Populate tables in pane
function populateLeadsTable(rows){
  const tbody = document.querySelector("#leadsTable tbody");
  tbody.innerHTML = "";
  rows.forEach(r => {
    const tr = document.createElement("tr");
    const cells = ["ID","Created On","Last Updated","Owner","Account","Name","Title/Role","Email","Phone","City/State","Source","Priority","Status","Stage","Est. Value ($)","Close Date","Notes"];
    cells.forEach(k => {
      const td = document.createElement("td");
      let val = r[k] ?? "";
      if (k==="Status"){
        const span = document.createElement("span"); span.className = "badge";
        span.textContent = val;
        if (val==="Won") span.classList.add("green");
        else if (val==="New"||val==="Proposal") span.classList.add("yellow");
        else if (val==="Lost") span.classList.add("red");
        td.appendChild(span);
      } else {
        td.textContent = val;
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}
async function refreshLeadsTable(){ const data = await getAllData(); populateLeadsTable(data.leads); }

function populateActivitiesTable(rows){
  const tbody = document.querySelector("#activitiesTable tbody");
  tbody.innerHTML = "";
  rows.forEach(r => {
    const tr = document.createElement("tr");
    ["Timestamp","Lead ID","Owner","Type","Notes","Next Step","Due Date"].forEach(k => {
      const td = document.createElement("td");
      td.textContent = r[k] ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

// Log activity prompt
async function logActivityPrompt(){
  const id = prompt("Enter Lead ID to attach activity:");
  if (!id) return;
  const type = prompt("Activity Type (Call/Email/Note/Meeting):","Call") || "";
  const notes = prompt("Notes:","") || "";
  const nextStep = prompt("Next Step:","") || "";
  const due = prompt("Due Date (yyyy-mm-dd):","") || "";
  await Excel.run(async (ctx)=>{
    const t = ctx.workbook.tables.getItem(TABLES.activities);
    const timestamp = new Date().toISOString().slice(0,16).replace("T"," ");
    t.rows.add(null, [[timestamp, Number(id)||id, "", type, notes, nextStep, due]]);
    await ctx.sync();
  });
  await refreshAll();
}

// ----- D3 CHARTS -----
function renderKPIs({leads}){
  const total = leads.length;
  const count = (key,val) => leads.filter(d=>String(d[key]||"")===val).length;
  document.getElementById("kpi-total").textContent = total;
  document.getElementById("kpi-new").textContent = count("Status","New");
  document.getElementById("kpi-qualified").textContent = count("Status","Qualified");
  document.getElementById("kpi-proposals").textContent = count("Status","Proposal");
  document.getElementById("kpi-won").textContent = count("Status","Won");
  document.getElementById("kpi-lost").textContent = count("Status","Lost");
}

function renderStatusChart(leads){
  const byStatus = d3.rollups(leads, v=>v.length, d=>d.Status||"").map(([k,v])=>({key:k||"Unknown", value:v}));
  drawPie("#chart-status", byStatus);
}

function renderStageValueChart(leads){
  const byStage = d3.rollups(leads, v=>d3.sum(v, d=>+d["Est. Value ($)"]||0), d=>d.Stage||"").map(([k,v])=>({key:k||"Unknown", value:v}));
  drawBars("#chart-stage", byStage);
}

function renderMonthChart(leads){
  const parse = d3.timeParse("%Y-%m-%d");
  const fmt = d3.timeFormat("%b %Y");
  const six = d3.timeMonth.offset(new Date(), -5);
  const months = d3.timeMonths(d3.timeMonth.floor(six), d3.timeMonth.offset(new Date(),1));
  const series = months.map(m => ({
    key: fmt(m),
    value: leads.filter(d => {
      const t = parse(d["Created On"]);
      if (!t) return false;
      return t.getFullYear()===m.getFullYear() && t.getMonth()===m.getMonth();
    }).length
  }));
  drawLine("#chart-months", series);
}

// Generic D3 components
function drawPie(sel, data){
  const svg = d3.select(sel); svg.selectAll("*").remove();
  const w = svg.node().clientWidth || 400, h = svg.node().clientHeight || 280, r = Math.min(w,h)/2 - 10;
  const g = svg.attr("viewBox", [-w/2, -h/2, w, h]).append("g");
  const arc = d3.arc().innerRadius(r*0.5).outerRadius(r);
  const pie = d3.pie().value(d=>d.value);
  const color = d3.scaleOrdinal().domain(data.map(d=>d.key)).range(d3.schemeTableau10);
  const arcs = g.selectAll("path").data(pie(data)).enter().append("path").attr("d", arc).attr("fill", d=>color(d.data.key)).append("title").text(d=>`${d.data.key}: ${d.data.value}`);
  // legend
  const legend = g.append("g").attr("transform",`translate(${r+20},${-r})`);
  const row = legend.selectAll("g").data(data).enter().append("g").attr("transform",(d,i)=>`translate(0,${i*18})`);
  row.append("rect").attr("width",12).attr("height",12).attr("fill", d=>color(d.key));
  row.append("text").attr("x",18).attr("y",10).text(d=>`${d.key} (${d.value})`).style("font-size","12px");
}

function drawBars(sel, data){
  const svg = d3.select(sel); svg.selectAll("*").remove();
  const w = svg.node().clientWidth || 400, h = svg.node().clientHeight || 280;
  const margin = {top:10,right:10,bottom:40,left:60};
  const x = d3.scaleBand().domain(data.map(d=>d.key)).range([margin.left, w-margin.right]).padding(0.2);
  const y = d3.scaleLinear().domain([0, d3.max(data,d=>d.value)||1]).nice().range([h-margin.bottom, margin.top]);
  const g = svg.attr("viewBox",[0,0,w,h]).append("g");
  g.append("g").attr("transform",`translate(0,${h-margin.bottom})`).call(d3.axisBottom(x)).selectAll("text").attr("transform","rotate(-20)").style("text-anchor","end");
  g.append("g").attr("transform",`translate(${margin.left},0)`).call(d3.axisLeft(y));
  g.selectAll("rect").data(data).enter().append("rect").attr("x",d=>x(d.key)).attr("y",d=>y(d.value)).attr("width",x.bandwidth()).attr("height",d=>y(0)-y(d.value)).attr("fill","#2563eb");
}

function drawLine(sel, data){
  const svg = d3.select(sel); svg.selectAll("*").remove();
  const w = svg.node().clientWidth || 800, h = svg.node().clientHeight || 360;
  const margin = {top:10,right:10,bottom:40,left:50};
  const x = d3.scalePoint().domain(data.map(d=>d.key)).range([margin.left, w-margin.right]);
  const y = d3.scaleLinear().domain([0, d3.max(data,d=>d.value)||1]).nice().range([h-margin.bottom, margin.top]);
  const g = svg.attr("viewBox",[0,0,w,h]).append("g");
  g.append("g").attr("transform",`translate(0,${h-margin.bottom})`).call(d3.axisBottom(x));
  g.append("g").attr("transform",`translate(${margin.left},0)`).call(d3.axisLeft(y));
  const line = d3.line().x(d=>x(d.key)).y(d=>y(d.value));
  g.append("path").datum(data).attr("fill","none").attr("stroke","#22c55e").attr("stroke-width",2).attr("d",line);
  g.selectAll("circle").data(data).enter().append("circle").attr("cx",d=>x(d.key)).attr("cy",d=>y(d.value)).attr("r",3).attr("fill","#22c55e");
}
