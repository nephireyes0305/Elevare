"use strict";

const ENDPOINT = "https://script.google.com/macros/s/AKfycbztrKqI5Xme2iRUF5L5Hfno3jHEWQvqFRlC3_WF7YgEL_7ieZyTmAWR5s0aTWw5cavs/exec";
const el = id => document.getElementById(id);
const status = el("status");

async function api(action, fields = {}) {
  const response = await fetch(ENDPOINT,{
    method:"POST",headers:{"Content-Type":"text/plain;charset=UTF-8"},
    body:JSON.stringify({action,...fields}),signal:AbortSignal.timeout(45000)
  });
  if (!response.ok) throw new Error("The registration service is unavailable.");
  const result = await response.json();
  if (!result.ok) throw new Error(result.error || "Unable to complete this request.");
  return result;
}

async function busy(button, message, operation) {
  button.disabled = true;
  message.textContent = "Please wait…";
  try { await operation(); }
  catch (error) { message.textContent = error.message || "Please retry."; }
  finally { button.disabled = false; }
}

function readReceipt(file) {
  if (!file || file.size > 2097152) throw new Error("Choose a receipt up to 2 MB.");
  return new Promise((resolve,reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({type:file.type,data:reader.result.split(",")[1]});
    reader.onerror = () => reject(new Error("Could not read the receipt."));
    reader.readAsDataURL(file);
  });
}

function registration() {
  const form = el("application");
  const submit = form.querySelector("[type=submit]");
  const chooser = form.elements.ambassadorId;
  const requestId = crypto.randomUUID();
  api("listLearnerAmbassadors").then(({ambassadors}) => {
    chooser.replaceChildren(new Option("Select your ambassador",""));
    for (const ambassador of ambassadors) chooser.add(new Option(ambassador.name,ambassador.id));
    submit.disabled = !ambassadors.length;
    if (!ambassadors.length) el("ambassador-note").textContent = "Registration will open when an ambassador becomes active.";
  }).catch(error => { status.textContent = error.message; });
  el("send-otp").onclick = () => {
    if (!form.elements.email.reportValidity()) return;
    busy(el("send-otp"),status,async () => {
      await api("sendLearnerOtp",{email:form.elements.email.value});
      status.textContent = "Check your inbox for your verification code.";
    });
  };
  form.onsubmit = event => {
    event.preventDefault();
    const data = new FormData(form);
    if (data.get("password") !== data.get("confirmPassword")) {
      status.textContent = "The passwords must match.";
      return;
    }
    const fields = Object.fromEntries(data);
    delete fields.confirmPassword;
    const names = ["firstName","middleName","lastName","suffix"];
    for (const key of names) fields[key] = data.get(key);
    fields.coLearners = [1,2,3].map((number) =>
      Object.fromEntries(names.map(key => [key,data.getAll(key)[number]])));
    busy(submit,status,async () => {
      fields.receipt = await readReceipt(data.get("receipt"));
      const result = await api("submitLearnerApplication",{...fields,requestId});
      form.reset();
      form.hidden = true;
      status.textContent = "Application received.\nLearner ID: " + result.learnerId +
        "\nStatus: Pending Verification." + (result.emailSent ? "" : "\nThe email could not be sent. Save this Learner ID.");
    });
  };
}

function authentication(isAdmin) {
  const form = el("sign-in");
  const key = isAdmin ? "elevare-review-session" : "elevare-learner-session";
  const panel = el(isAdmin ? "admin-portal" : "learner-portal");
  let token = sessionStorage.getItem(key);
  function sessionReady(result) {
    status.textContent = "";
    form.hidden = true;
    panel.hidden = false;
    if (!isAdmin) el("welcome").textContent = "Welcome, " + (result.member.name || result.member.username) + ".";
  }
  form.onsubmit = event => {
    event.preventDefault();
    const credentials = Object.fromEntries(new FormData(form));
    busy(event.submitter,status,async () => {
      const result = await api(isAdmin ? "adminLogin" : "login",{...credentials});
      token = result.token;
      sessionStorage.setItem(key,token);
      form.reset();
      sessionReady(result);
    });
  };
  if (token) api(isAdmin ? "adminSession" : "getSession",{token}).then(sessionReady)
    .catch(() => { sessionStorage.removeItem(key); token = null; });
  el(isAdmin ? "sign-out" : "learner-sign-out").onclick = async () => {
    try { if (token) await api(isAdmin ? "adminLogout" : "logout",{token}); }
    finally { sessionStorage.removeItem(key); location.reload(); }
  };
  if (isAdmin) registrationReview(() => token);
}

function details(node, fields) {
  const list = document.createElement("dl");
  for (const [label,value] of fields) {
    const term = document.createElement("dt");
    const data = document.createElement("dd");
    term.textContent = label;
    data.textContent = value || "—";
    list.append(term,data);
  }
  node.replaceChildren(list);
}

function registrationReview(getToken) {
  const dialog = el("review-window");
  const message = el("review-status");
  const tabs = [...dialog.querySelectorAll("[data-section]")];
  const decisions = [...dialog.querySelectorAll("[data-decision]")];
  let current = null;
  let receiptUrl = "";
  const name = person => [person.firstName,person.middleName,person.lastName,person.suffix].filter(Boolean).join(" ");

  function tab(selected) {
    for (const button of tabs) {
      const active = button.dataset.section === selected;
      button.setAttribute("aria-pressed",active);
      el(button.dataset.section).hidden = !active;
    }
  }
  for (const button of tabs) button.onclick = () => tab(button.dataset.section);
  el("open-review").onclick = () => dialog.showModal();
  dialog.addEventListener("close",() => {
    if (receiptUrl) URL.revokeObjectURL(receiptUrl);
    receiptUrl = "";
    current = null;
    el("review").hidden = true;
  });

  function render(application) {
    current = application;
    const p = current.personal;
    details(el("personal"),[["Learner ID",current.learnerId],["Status",current.status],
      ["Name",name(p)],["Username",current.username],["Email (verified by OTP)",current.email],
      ["Birthdate",p.birthdate],["Gender",p.gender],["Marital status",p.maritalStatus],
      ["Phone",p.phone],["Additional phone",p.phone2],["Address",p.address],["Ambassador",p.ambassadorName]]);
    details(el("co"),current.coLearners.map(person => [person.id,name(person)]));
    details(el("payment"),[["Payment method",current.method],["Reference number",current.reference]]);
    if (receiptUrl) URL.revokeObjectURL(receiptUrl);
    const receipt = current.receipt;
    receiptUrl = URL.createObjectURL(new Blob([Uint8Array.from(atob(receipt.data),c=>c.charCodeAt(0))],{type:receipt.type}));
    if (receipt.type.startsWith("image/")) {
      const image = new Image();
      image.alt = "Submitted proof of payment";
      image.src = receiptUrl;
      image.className = "receipt";
      el("payment").append(image);
    }
    const link = document.createElement("a");
    link.href = receiptUrl;
    link.target = "_blank";
    link.rel = "noopener";
    link.textContent = "Open receipt";
    el("payment").append(link);
    el("decision-comment").value = current.comment || "";
    for (const button of decisions) {
      const pending = current.status === "Pending Verification";
      const retry = current.status === button.dataset.decision && current.emailStatus !== "Decision email sent";
      const correction = button.dataset.decision === "Correction Requested";
      const resend = correction && current.status === "Correction Requested";
      button.hidden = !pending && !retry && !resend;
      button.textContent = correction ? (resend ? "Resend correction link" : "Request Correction") : retry ? "Retry decision email" : button.dataset.decision === "Verified" ? "Verify" : "Reject";
    }
    el("decision-comment").disabled = current.status !== "Pending Verification";
    el("review").hidden = false;
    tab("personal");
  }

  el("lookup").onsubmit = event => {
    event.preventDefault();
    el("review").hidden = true;
    const learnerId = el("lookup").elements.learnerId.value.trim();
    busy(event.submitter,message,async () => {
      const result = await api("adminLearnerApplication",{token:getToken(),learnerId});
      if (!dialog.open) return;
      message.textContent = "";
      render(result.application);
    });
  };
  for (const button of decisions) button.onclick = () => {
    if (!current) return;
    const comment = el("decision-comment").value.trim();
    if (button.dataset.decision === "Correction Requested" && current.status === "Pending Verification" && !comment) {
      message.textContent = "Explain what the learner should correct.";
      el("decision-comment").focus();
      return;
    }
    if (!confirm("Submit this registration decision?")) return;
    decisions.forEach(item => { item.disabled = true; });
    busy(button,message,async () => {
      const result = await api("adminDecideLearner",{
        token:getToken(),learnerId:current.learnerId,decision:button.dataset.decision,
        comment:el("decision-comment").value
      });
      current.comment = comment;
      current.status = result.status;
      current.emailStatus = result.emailSent ? "Decision email sent" : "Decision email failed";
      render(current);
      message.textContent = result.status + (result.emailSent ? ". Email sent." : ". Email could not be sent. Retry the decision email.");
    }).finally(() => decisions.forEach(item => { item.disabled = false; }));
  };
}

function correctionForm() {
  const form = el("correct-payment");
  const params = new URLSearchParams(location.hash.slice(1));
  const learnerId = params.get("id");
  const token = params.get("token");
  if (!learnerId || !token) {
    form.hidden = true;
    status.textContent = "Open the private correction link in your email.";
    return;
  }
  form.onsubmit = event => {
    event.preventDefault();
    const fields = Object.fromEntries(new FormData(form));
    busy(event.submitter,status,async () => {
      fields.receipt = await readReceipt(fields.receipt);
      await api("resubmitLearnerPayment",{...fields,learnerId,token});
      form.reset();
      form.hidden = true;
      history.replaceState(null,"",location.pathname);
      status.textContent = "Payment details received. Your application is Pending Verification.";
    });
  };
}

const page = document.body.dataset.page;
if (page === "register") registration();
else if (page === "correction") correctionForm();
else authentication(page === "admin");
