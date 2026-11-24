"use strict";

(() => {
  const get = id => document.getElementById(id);
  const portal = get("admin-portal");
  const status = get("dashboard-status");
  const time = new Intl.DateTimeFormat("en-PH",{
    dateStyle:"medium",timeStyle:"short",timeZone:"Asia/Manila"
  });
  const formatDate = timestamp => Number.isFinite(timestamp) ? time.format(timestamp) : "—";
  const order = {"Pending":0,"On Going":1,"Not Started":2};
  let records = [];
  let ambassadors = [];
  let existingLearners = [];
  let selectedAmbassador = "";
  let exportRows = null;

  function button(label, action) {
    const item = document.createElement("button");
    item.type = "button";
    item.textContent = label;
    item.addEventListener("click",action);
    return item;
  }

  function table(node, headers, rows, empty) {
    node.replaceChildren();
    if (!rows.length) {
      const message = document.createElement("p");
      message.textContent = empty;
      node.append(message);
      return;
    }
    const wrapper = document.createElement("div");
    wrapper.className = "admin-table";
    const grid = document.createElement("table");
    const head = grid.createTHead().insertRow();
    for (const label of headers) {
      const cell = document.createElement("th");
      cell.scope = "col";
      cell.textContent = label;
      head.append(cell);
    }
    const body = grid.createTBody();
    for (const values of rows) {
      const row = body.insertRow();
      for (const value of values) {
        const cell = row.insertCell();
        if (value instanceof HTMLElement) cell.append(value);
        else cell.textContent = value ?? "—";
      }
    }
    wrapper.append(grid);
    node.append(wrapper);
  }

  function openApplication(id) {
    get("open-review").click();
    const form = get("lookup");
    form.elements.learnerId.value = id;
    form.requestSubmit(form.querySelector("button"));
  }

  function renderLearners() {
    const sort = get("learner-sort").value;
    const recent = (a,b) => (b.submittedAt ?? -Infinity) - (a.submittedAt ?? -Infinity);
    const rows = [...records,...existingLearners].sort((a,b) => {
      if (sort === "status") return (order[a.trainingStatus] ?? 9) - (order[b.trainingStatus] ?? 9) || recent(a,b);
      if (sort === "oldest") return (a.submittedAt ?? Infinity) - (b.submittedAt ?? Infinity);
      return recent(a,b);
    });
    table(get("learner-list"),["Learner","ID","Enrollment date","Application","Training status","Completed"],rows.map(r =>
      [r.name,r.id,formatDate(r.submittedAt),r.status,r.trainingStatus,r.completedTrainings + " of 2"]
    ),"No learners found.");
  }

  function renderReferrals() {
    const target = get("referral-list");
    const ambassador = ambassadors.find(item => item.id === selectedAmbassador);
    target.hidden = !ambassador;
    if (!ambassador) return;
    const title = document.createElement("h3");
    title.textContent = ambassador.name + " · Referrals";
    const list = document.createElement("div");
    table(list,["Learner","Learner ID","Registered","Application status"],
      records.filter(r => r.ambassadorId === ambassador.id).map(r =>
        [r.name,r.id,formatDate(r.submittedAt),r.status]),"No referrals for this ambassador.");
    target.replaceChildren(title,list);
  }

  function render() {
    const pending = records.filter(r => r.status === "Pending Verification")
      .sort((a,b) => a.submittedAt - b.submittedAt);
    get("pending-count").textContent = "(" + pending.length + ")";
    table(get("pending-list"),["Learner","Learner ID","Submitted","Ambassador","Action"],pending.map(r =>
      [r.name,r.id,formatDate(r.submittedAt),r.ambassador || "—",button("Review",() => openApplication(r.id))]
    ),"No pending applications.");
    renderLearners();
    table(get("ambassador-list"),["Ambassador","ID","Status","Referrals","Action"],ambassadors.map(item => [
      item.name,item.id,item.status,records.filter(r => r.ambassadorId === item.id).length,
      button("View referrals",() => { selectedAmbassador = item.id; renderReferrals(); })
    ]),"No ambassadors found.");
    renderReferrals();
  }

  async function refresh() {
    const token = sessionStorage.getItem("elevare-review-session");
    if (!token) return;
    get("reload-dashboard").disabled = true;
    status.textContent = "Loading dashboard…";
    try {
      const response = await api("adminElevareDashboard",{token});
      records = response.applications;
      ambassadors = response.ambassadors;
      existingLearners = response.existingLearners;
      render();
      status.textContent = "";
    } catch (error) {
      status.textContent = error.message;
    } finally {
      get("reload-dashboard").disabled = false;
    }
  }

  function generate(event) {
    event.preventDefault();
    const start = Date.parse(event.target.elements.start.value + ":00+08:00");
    const end = Date.parse(event.target.elements.end.value + ":59.999+08:00");
    const output = get("report-output");
    const summary = get("report-status");
    output.replaceChildren();
    exportRows = null;
    get("export-report").disabled = true;
    if (!Number.isFinite(start) || !Number.isFinite(end) || start > end) {
      summary.textContent = "Enter a valid start and end date and time.";
      return;
    }
    const groups = new Map();
    const matches = records.filter(r => Number.isFinite(r.submittedAt) && r.submittedAt >= start && r.submittedAt <= end)
      .sort((a,b) => a.submittedAt - b.submittedAt);
    for (const record of matches) {
      const key = record.ambassadorId || record.ambassador || "";
      if (!groups.has(key)) groups.set(key,[]);
      groups.get(key).push(record);
    }
    summary.textContent = matches.length + " registrations · " + formatDate(start) + " through " + formatDate(end);
    exportRows = [["Ambassador","Total registrations","Learner","Learner ID","Registration time (Philippines)","Application status"]];
    for (const entries of [...groups.values()].sort((a,b) => (a[0].ambassador || "").localeCompare(b[0].ambassador || ""))) {
      const name = entries[0].ambassador || "Unassigned";
      const heading = document.createElement("h3");
      heading.textContent = name + " · " + entries.length;
      const list = document.createElement("div");
      const rows = entries.map(r => [r.name,r.id,formatDate(r.submittedAt),r.status]);
      table(list,["Learner","Learner ID","Registered","Application status"],rows,"No registrations.");
      output.append(heading,list);
      exportRows.push(...rows.map(row => [name,entries.length,...row]));
    }
    get("export-report").disabled = !matches.length;
  }

  function downloadReport() {
    if (!exportRows) return;
    const csv = exportRows.map(row => row.map(value => {
      const text = String(value ?? "");
      const safe = /^[\s]*[=+@-]|^[\t\r\n]/.test(text) ? "'" + text : text;
      return '"' + safe.replaceAll('"','""') + '"';
    }).join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob(["\uFEFF",csv],{type:"text/csv;charset=utf-8"}));
    const link = document.createElement("a");
    link.href = url;
    link.download = "elevare-ambassador-referrals.csv";
    link.click();
    setTimeout(() => URL.revokeObjectURL(url),1000);
  }

  const tabs = [...portal.querySelectorAll("[data-view]")];
  for (const tab of tabs) {
    tab.addEventListener("click",() => {
      for (const item of tabs) {
        const active = item === tab;
        item.setAttribute("aria-pressed",String(active));
        get("view-" + item.dataset.view).hidden = !active;
      }
    });
  }
  get("learner-sort").addEventListener("change",renderLearners);
  get("reload-dashboard").addEventListener("click",refresh);
  get("ambassador-placeholder").addEventListener("click",() => get("ambassador-dialog").showModal());
  get("report-range").addEventListener("submit",generate);
  get("export-report").addEventListener("click",downloadReport);
  get("review-window").addEventListener("close",refresh);
  new MutationObserver(() => { if (!portal.hidden) refresh(); })
    .observe(portal,{attributes:true,attributeFilter:["hidden"]});
  if (!portal.hidden) refresh();
})();
