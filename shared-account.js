const api = "https://script.google.com/macros/s/AKfycbztrKqI5Xme2iRUF5L5Hfno3jHEWQvqFRlC3_WF7YgEL_7ieZyTmAWR5s0aTWw5cavs/exec";
const byId = (id) => document.getElementById(id);
const value = (id) => byId(id).value.trim();

async function request(action, fields) {
  const response = await fetch(api, {
    method: "POST",
    headers: {"Content-Type":"text/plain;charset=UTF-8"},
    body: JSON.stringify({action, ...fields}),
    signal: AbortSignal.timeout(20000)
  });
  if (!response.ok) throw new Error("The account service is unavailable.");
  const result = await response.json();
  if (!result.ok) throw new Error(result.error || "Unable to complete this step.");
  return result;
}

function bind(id, handler) {
  const button = byId(id);
  if (!button) return;
  button.addEventListener("click", async () => {
    button.disabled = true;
    try { await handler(); }
    catch (error) { alert(error.message); }
    finally { button.disabled = false; }
  });
}

if (byId("verifyCodeBtn")) {
  let step = 1;
  let code = "";
  function go(next) {
    step = next;
    document.querySelectorAll(".reg-panel").forEach(panel =>
      panel.classList.toggle("active", Number(panel.dataset.step) === next));
    document.querySelectorAll(".step").forEach((item, index) =>
      item.classList.toggle("active", index === Math.min(next, 4) - 1));
    window.scrollTo(0, 0);
  }
  document.querySelectorAll(".back").forEach(button => {
    button.onclick = () => go(step - 1);
  });
  bind("verifyCodeBtn", async () => {
    const entered = value("regCode").toUpperCase();
    await request("verifyElevareCode", {code:entered});
    code = entered;
    go(2);
  });
  document.querySelector(".next-details").onclick = () => {
    const fields = [...document.querySelectorAll('[data-step="2"] input, [data-step="2"] select')];
    if (!fields.every(field => field.reportValidity())) return;
    if (byId("rPassword").value.length < 8) return alert("Use at least eight characters for your password.");
    byId("emailDisplay").value = value("rEmail");
    byId("emailOtp").value = "";
    go(3);
  };
  const verify = byId("verifyEmailOtp");
  verify.hidden = true;
  bind("sendEmailOtp", async () => {
    await request("sendEmailOtp", {email:value("rEmail")});
    alert("Email verification code sent.");
  });
  bind("toDependents", () => {
    if (!/^\d{6}$/.test(value("emailOtp"))) return alert("Enter the six-digit code sent to your email.");
    go(4);
  });
  bind("addRegDependent", () => {
    const row = document.createElement("fieldset");
    row.className = "dependent-box";
    row.innerHTML = '<legend>Dependent</legend><label>Full name<input class="dname" required></label><label>Relationship<select class="drel"><option>Spouse</option><option>Life Partner</option><option>Child</option><option>Parent</option><option>Sibling</option></select></label><label>Birthdate<input class="dbirth" type="date" required></label><label>Gender<select class="dgender"><option>Male</option><option>Female</option><option>Other</option></select></label><button type="button">Remove</button>';
    row.querySelector("button").onclick = () => row.remove();
    byId("regDependents").append(row);
  });
  bind("finishRegistration", async () => {
    const rows = [...document.querySelectorAll(".dependent-box")];
    if (!rows.every(row => [...row.querySelectorAll("input")].every(input => input.reportValidity()))) return;
    const dependents = rows.map(row => ({
      name:row.querySelector(".dname").value.trim(),
      relationship:row.querySelector(".drel").value,
      birthdate:row.querySelector(".dbirth").value,
      gender:row.querySelector(".dgender").value
    }));
    const result = await request("registerElevareMember", {
      registrationCode:code, otp:value("emailOtp"), username:value("rUsername"),
      password:byId("rPassword").value, name:value("rName"), address:value("rAddress"),
      phone:value("rPhone"), phone2:value("rPhone2"), email:value("rEmail"),
      birthdate:value("rBirthdate"), gender:value("rGender"),
      maritalStatus:value("rMaritalStatus"), dependents
    });
    byId("rPassword").value = "";
    byId("emailOtp").value = "";
    byId("registrationSummary").textContent = "Account ID: " + result.memberId;
    go(5);
  });
}

const login = byId("shared-login");
if (login) {
  const status = byId("login-status");
  const sessionKey = "elevare-session";
  function showMember(member) {
    login.hidden = true;
    byId("signed-in").hidden = false;
    status.textContent = "Signed in as " + member.name + ".";
  }
  login.onsubmit = async (event) => {
    event.preventDefault();
    const button = login.querySelector("button");
    button.disabled = true;
    try {
      const result = await request("loginElevareMember", {
        username:login.elements.username.value.trim(), password:login.elements.password.value
      });
      sessionStorage.setItem(sessionKey, result.token);
      login.reset();
      showMember(result.member);
    } catch (error) { status.textContent = error.message; }
    finally { button.disabled = false; }
  };
  const token = sessionStorage.getItem(sessionKey);
  if (token) request("getElevareSession", {token}).then(result => showMember(result.member))
    .catch(() => sessionStorage.removeItem(sessionKey));
  byId("elevare-signout").onclick = () => {
    sessionStorage.removeItem(sessionKey);
    location.reload();
  };
}
