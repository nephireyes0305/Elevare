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
codeForm?.addEventListener("submit", (event) => {
  event.preventDefault();
  codeForm.querySelector(".form-status").textContent =
    "Code verification is unavailable until Elevare connects to the CareLink registration service.";
});

document.getElementById("enrollment-form")?.addEventListener("submit", (event) => {
  event.preventDefault();
});
