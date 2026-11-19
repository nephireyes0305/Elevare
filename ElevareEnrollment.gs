function completeElevareRegistration(body) {
  const code = clean(body.code).toUpperCase();
  const name = clean(body.name);
  const email = clean(body.email).toLowerCase();
  const phone = clean(body.phone);
  if (name.length < 2 || name.length > 120 ||
      email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
      !/^\+?[0-9 ()-]{7,24}$/.test(phone) || phone.replace(/\D/g, "").length < 7) {
    return {ok:false,error:"Enter a valid full name, email address, and mobile number."};
  }

  // This helper takes its own lock while it initializes the status column.
  const statusColumn = setupElevareRegistrationStatus();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SS.getSheetByName("Registration Codes");
    const rows = sheet.getDataRange().getDisplayValues();
    const matchingRows = [];
    for (let i = 1; i < rows.length; i++) {
      if (clean(rows[i][0]).toUpperCase() === code) matchingRows.push(i + 1);
    }
    const denied = {ok:false,error:"This code is unavailable for Elevare registration."};
    if (!/^[A-Z0-9-]{8,100}$/.test(code) || matchingRows.length !== 1) return denied;

    const rowNumber = matchingRows[0];
    const row = rows[rowNumber - 1];
    const status = clean(row[statusColumn - 1]).toLowerCase();
    let detailsColumn = headerMap(sheet).elevareenrollment;
    if (status === "used" && detailsColumn) {
      const saved = JSON.parse(row[detailsColumn - 1] || "null");
      if (saved && saved.name === name && saved.email === email && saved.phone === phone) {
        return {ok:true};
      }
    }
    if (clean(row[2]).toLowerCase() !== "active" || status !== "paid") return denied;

    if (!detailsColumn) {
      detailsColumn = sheet.getLastColumn() + 1;
      if (detailsColumn > sheet.getMaxColumns()) sheet.insertColumnAfter(detailsColumn - 1);
      sheet.getRange(1, detailsColumn).setValue("Elevare Enrollment");
    }
    sheet.getRange(rowNumber, detailsColumn).setValue(JSON.stringify({
      name, email, phone, registeredAt: new Date().toISOString()
    }));
    SpreadsheetApp.flush();
    sheet.getRange(rowNumber, statusColumn).setValue("Used");
    SpreadsheetApp.flush();
    return {ok:true};
  } finally {
    lock.releaseLock();
  }
}
