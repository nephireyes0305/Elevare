# Shared Elevare and CareLink registration

1. Back up the current CareLink Apps Script. This Code.gs is based on the script you shared. If the live project has other newer changes, merge those before deployment. If the previous ElevareEnrollment.gs was added as a separate file, it can stay, but the new site does not call its old completion action.
2. In Apps Script Project Settings, add script property CARELINK_SPREADSHEET_ID with the ID of the same private spreadsheet your current script uses. Alternatively, keep the original const SS line from your current script in place of the packaged one.
3. Update the existing Apps Script with Code.gs and save. If the previous setup put setupElevareRegistrationStatus or verifyElevareCode in separate files, remove the duplicate definitions so each function exists once. Update the existing web app deployment to a new version, keeping its URL.
4. Upload only register.html, shared-registration.html, shared-account.js and login.html to the root of the Elevare GitHub repository. Code.gs goes in Apps Script.
5. Test on a copy first. Use a fresh disposable Active code with Elevare Status Paid. Complete the Elevare form with email OTP and verify that a single Members row is created with Elevare Only status and a single set of credentials. Sign in on Elevare. After CareLink payment and donation have been verified, submit the existing CareLink registration form with the same code and the same username, password and email. Confirm the original Members row is activated, the CareLink Used flag is checked, and the same credentials sign in on both sites.

The original CareLink signup form is reused on Elevare. CareLink's existing form activates the saved account and does not create a duplicate. Optional dependents saved on Elevare are retained. The Elevare login screen shows the public catalog link after authentication. Private course content is not included.

Registrations completed through the earlier three-field intake have no shared credentials. Staff should reconcile those rows separately before reissuing any code. Do not reset codes already used to activate CareLink membership.
