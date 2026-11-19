# Install the Elevare registration form

1. In your current CareLink Apps Script project, add a script named ElevareEnrollment and paste ElevareEnrollment.gs. Keep the existing verification and CareLink registration rules.
2. In the existing doPost switch, immediately after the verifyElevareCode case, add:

   ```javascript
   case "completeElevareRegistration": return json(completeElevareRegistration(body));
   ```

3. Save and publish a new version of the existing web app deployment, keeping the same URL.
4. Replace register.html and portal.js in the Elevare GitHub repository. Wait for GitHub Pages to publish.
5. Test with a disposable Active code whose Elevare Status is Paid. Verify it, complete the form with test details, and confirm the same row shows Elevare Status Used with details in Elevare Enrollment. This leaves the CareLink Used checkbox untouched.

The form records Elevare enrollment. CareLink account creation, login, and private course access are separate from this update.
