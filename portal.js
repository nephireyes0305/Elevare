document.querySelectorAll("[data-dialog]").forEach((button) => {
  button.addEventListener("click", () => {
    const dialog = document.getElementById(button.dataset.dialog);
    if (button.dataset.course) {
      document.getElementById("enrollment-course").textContent = button.dataset.course;
    }
    dialog.showModal();
  });
});

document.querySelectorAll("[data-close]").forEach((button) => {
  button.addEventListener("click", () => button.closest("dialog").close());
});

const codeForm = document.getElementById("code-form");
const registrationForm = document.getElementById("elevare-registration");
let verifiedCode = "";

async function callCareLink(request) {
  const url = "https://script.google.com/macros/s/AKfycbztrKqI5Xme2iRUF5L5Hfno3jHEWQvqFRlC3_WF7YgEL_7ieZyTmAWR5s0aTWw5cavs/exec";
  const response = await fetch(url, {
    method: "POST",
    headers: {"Content-Type": "text/plain;charset=UTF-8"},
    body: JSON.stringify(request),
    signal: AbortSignal.timeout(20000)
  });
  if (!response.ok) throw new Error("Request failed");
  const result = await response.json();
  if (result.stage === "routing" || result.stage === "server") {
    throw new Error("Service unavailable");
  }
  return result;
}

async function submitRequest(form, request, onSuccess) {
  const button = form.querySelector('button[type="submit"]');
  const status = form.querySelector(".form-status");
  button.disabled = true;
  status.textContent = "Please wait…";
  try {
    const result = await callCareLink(request);
    status.textContent = result.ok ? "" : (result.error || "Unable to complete the request.");
    if (result.ok) onSuccess();
  } catch {
    status.textContent = "The registration service is unavailable right now. Please try again.";
  } finally {
    button.disabled = false;
  }
}

codeForm?.addEventListener("submit", (event) => {
  event.preventDefault();
  const code = codeForm.elements.code.value.trim().toUpperCase();
  submitRequest(codeForm, {action:"verifyElevareCode", code}, () => {
    if (!registrationForm) return;
    verifiedCode = code;
    codeForm.hidden = true;
    registrationForm.hidden = false;
    registrationForm.elements.fullName.focus();
  });
});

registrationForm?.addEventListener("submit", (event) => {
  event.preventDefault();
  const fields = new FormData(registrationForm);
  submitRequest(registrationForm, {
    action:"completeElevareRegistration",
    code:verifiedCode,
    name:fields.get("fullName"),
    email:fields.get("email"),
    phone:fields.get("phone")
  }, () => {
    verifiedCode = "";
    codeForm.reset();
    registrationForm.reset();
    registrationForm.hidden = true;
    const message = document.getElementById("registration-complete");
    message.hidden = false;
    message.textContent = "Elevare registration complete. You can now use the same code on CareLink if its registration requirements are met.";
  });
});

document.getElementById("enrollment-form")?.addEventListener("submit", (event) => {
  event.preventDefault();
});
