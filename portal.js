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

const careLinkApi = "https://script.google.com/macros/s/AKfycbztrKqI5Xme2iRUF5L5Hfno3jHEWQvqFRlC3_WF7YgEL_7ieZyTmAWR5s0aTWw5cavs/exec";
const codeForm = document.getElementById("code-form");
codeForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const code = codeForm.elements.code.value.trim().toUpperCase();
  const status = codeForm.querySelector(".form-status");
  const button = codeForm.querySelector('button[type="submit"]');

  if (!/^[A-Z0-9-]{8,100}$/.test(code)) {
    status.textContent = "Enter the Elevare registration code provided by your ambassador.";
    return;
  }

  button.disabled = true;
  status.textContent = "Checking your registration code…";
  try {
    const response = await fetch(careLinkApi, {
      method: "POST",
      headers: {"Content-Type": "text/plain;charset=UTF-8"},
      body: JSON.stringify({action: "verifyElevareCode", code}),
      signal: AbortSignal.timeout(15000)
    });
    if (!response.ok) throw new Error("Request failed");
    const result = await response.json();
    if (result.stage === "routing" || result.stage === "server") {
      throw new Error("Verification service unavailable");
    }
    status.textContent = result.ok === true
      ? "Elevare code verified. Account registration is pending setup. Your code has not been used."
      : "This code is unavailable for Elevare registration. Check with your ambassador.";
  } catch {
    status.textContent = "The verification service is unavailable right now. Please try again later.";
  } finally {
    button.disabled = false;
  }
});

document.getElementById("enrollment-form")?.addEventListener("submit", (event) => {
  event.preventDefault();
});
