// Restore the original const SS = ... line from your existing Apps Script here.
const SS = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("CARELINK_SPREADSHEET_ID"));
const SESSION_SECONDS = 21600; // 6 hours

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return json({ok:true,service:"Carelink Member System API",status:"online",version:"9.1",timestamp:new Date().toISOString()});
}
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return json({ok:false,stage:"request",error:"No POST body received."});
    let body;
    try { body=JSON.parse(e.postData.contents); }
    catch(err) { return json({ok:false,stage:"parse",error:"Invalid JSON: "+err.message}); }

    switch(body.action) {
      case "validateCode": return json(validateRegistrationCode(body.code || body.registrationCode));
      case "sendEmailOtp": return json(sendEmailOtp(body.email));
      case "verifyEmailOtp": return json(verifyEmailOtp(body.email,body.otp));
      case "registerMember": return json(registerMember(body));
      case "registerElevareMember": return json(registerMember(body,true));
      case "loginElevareMember": return json(loginMember(body.username,body.password,true));
      case "getElevareSession": return json(getSession(body.token,true));
      case "login": return json(loginMember(body.username,body.password));
      case "getSession": return json(getSession(body.token));
      case "getMemberPortalData": return json(getMemberPortalData(body.token));
      case "downloadStatement": return json(downloadContributionStatement(body.token));
      case "requestAssistance": return json(requestAssistance(body));
      case "addPortalDependent": return json(addPortalDependent(body));
      case "logout": return json(logoutMember(body.token));
      case "addDependent": return json(addDependent(body.memberId,body.dependent));
      case "adminLogin": return json(loginAdmin(body.username,body.password));
      case "adminSession": return json(getAdminSession(body.token));
      case "adminLogout": return json(adminLogout(body.token));
      case "adminDashboard": return json(adminDashboard(body));
      case "adminListMembers": return json(adminListMembers(body));
      case "adminUpdateMember": return json(adminUpdateMember(body));
      case "adminListCodes": return json(adminListCodes(body));
      case "adminSaveCode": return json(adminSaveCode(body));
      case "adminListContributions": return json(adminListContributions(body));
      case "adminSaveContribution": return json(adminSaveContribution(body));
      case "adminListAssistance": return json(adminListAssistance(body));
      case "adminUpdateAssistance": return json(adminUpdateAssistance(body));
      case "adminAddFundTransaction": return json(adminAddFundTransaction(body));
      case "verifyElevareCode": return json(verifyElevareCode(body.code));
      default: return json({ok:false,stage:"routing",error:"Unknown action: "+body.action});
    }
  } catch(err) {
    console.error(err.stack || err);
    return json({ok:false,stage:"server",error:String(err.message || err)});
  }
}

function boolVal(v){ return v === true || String(v).trim().toLowerCase()==="true"; }
function clean(v){ return String(v == null ? "" : v).trim(); }
function addDays(d,n){ return new Date(d.getTime()+n*86400000); }


// v8.1: table-safe database writer.
function firstEmptyDataRow(sh,keyColumn){
  const start=2, last=Math.max(sh.getLastRow(),start);
  const vals=sh.getRange(start,keyColumn,last-start+1,1).getDisplayValues();
  for(let i=0;i<vals.length;i++) if(!String(vals[i][0]).trim()) return start+i;
  return last+1;
}
function writeDatabaseRow(sh,values,keyColumn){
  keyColumn=keyColumn||1;
  const row=firstEmptyDataRow(sh,keyColumn);
  if(sh.getMaxColumns()<values.length) sh.insertColumnsAfter(sh.getMaxColumns(),values.length-sh.getMaxColumns());
  sh.getRange(row,1,1,values.length).setValues([values]);
  SpreadsheetApp.flush();
  if(String(sh.getRange(row,keyColumn).getDisplayValue()).trim()!==String(values[keyColumn-1]).trim())
    throw new Error("Database write verification failed on "+sh.getName()+" row "+row);
  return row;
}
function audit(action,memberId,detail,status){
  try{
    const sh=SS.getSheetByName("Audit Log"); if(!sh) return;
    writeDatabaseRow(sh,["AUD-"+Utilities.getUuid().slice(0,8).toUpperCase(),new Date(),action,memberId||"",detail||"",status||""],1);
  }catch(e){ console.error("Audit error: "+e.message); }
}

function validateRegistrationCode(code) {
  code=clean(code).toUpperCase();
  const sh=SS.getSheetByName("Registration Codes");
  if(!sh) return {ok:false,error:'Sheet "Registration Codes" not found.'};
  const v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++){
    if(clean(v[i][0]).toUpperCase()===code){
      if(clean(v[i][2]).toLowerCase()!=="active") return {ok:false,error:"Registration code is not active."};
      if(boolVal(v[i][3])) return {ok:false,error:"Registration code has already been used."};
      if(!boolVal(v[i][6]) || !boolVal(v[i][7])) return {ok:false,error:"Registration payment and initial donation must be confirmed."};
      const elevareColumn=headerMap(sh).elevarestatus;
      if(!elevareColumn || clean(v[i][elevareColumn-1]).toLowerCase()!=="used")
        return {ok:false,error:"Complete Elevare registration before registering with CareLink."};
      return {ok:true,code:code};
    }
  }
  return {ok:false,error:"Registration code not found."};
}

function sendEmailOtp(email) {
  try {
    email=clean(email).toLowerCase();
    if(!email || email.indexOf("@")<1) return {ok:false,stage:"validation",error:"Invalid email address."};
    const sh=SS.getSheetByName("OTP Logs");
    if(!sh) return {ok:false,stage:"sheet",error:'Sheet "OTP Logs" was not found.'};

    const otp=String(Math.floor(100000+Math.random()*900000));
    const now=new Date(), expires=new Date(now.getTime()+10*60*1000);
    const otpId="OTP-"+Utilities.getUuid().slice(0,8).toUpperCase();
    writeDatabaseRow(sh,[otpId,"","Email",email,otp,now,expires,"","Pending"],1);

    try {
      MailApp.sendEmail({
        to:email,
        subject:"Carelink Email Verification Code",
        htmlBody:'<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto"><h2 style="color:#164b55">Carelink Email Verification</h2><p>Use this verification code to continue your Carelink registration:</p><div style="font-size:30px;font-weight:bold;letter-spacing:8px;color:#164b55;margin:24px 0">'+otp+'</div><p>This code expires in <b>10 minutes</b> and can only be used once.</p><p>If you did not request this code, you may ignore this email.</p><hr><small>Carelink Community Welfare Services Inc.</small></div>'
      });
    } catch(err) {
      try { sh.getRange(sh.getLastRow(),9).setValue("Send Failed"); } catch(_) {}
      return {ok:false,stage:"email-send",error:err.message};
    }
    return {ok:true,message:"Verification email sent."};
  } catch(err) { return {ok:false,stage:"sendEmailOtp",error:String(err.message||err)}; }
}

function verifyEmailOtp(email,otp) {
  email=clean(email).toLowerCase(); otp=clean(otp);
  const sh=SS.getSheetByName("OTP Logs");
  if(!sh) return {ok:false,error:'Sheet "OTP Logs" was not found.'};
  const rows=sh.getDataRange().getValues(), now=new Date();
  for(let i=rows.length-1;i>=1;i--){
    if(clean(rows[i][2])==="Email" && clean(rows[i][3]).toLowerCase()===email && clean(rows[i][8])==="Pending"){
      const expires=rows[i][6] instanceof Date ? rows[i][6] : new Date(rows[i][6]);
      if(now>expires){ sh.getRange(i+1,9).setValue("Expired"); return {ok:false,error:"OTP has expired. Please request a new code."}; }
      if(clean(rows[i][4])!==otp) return {ok:false,error:"Incorrect verification code."};
      sh.getRange(i+1,8).setValue(now); sh.getRange(i+1,9).setValue("Verified");
      return {ok:true};
    }
  }
  return {ok:false,error:"No active OTP found. Please request a new code."};
}
function emailWasVerified(email) {
  const sh=SS.getSheetByName("OTP Logs"); if(!sh) return false;
  const rows=sh.getDataRange().getValues(); email=clean(email).toLowerCase();
  for(let i=rows.length-1;i>=1;i--)
    if(clean(rows[i][2])==="Email" && clean(rows[i][3]).toLowerCase()===email && clean(rows[i][8])==="Verified") return true;
  return false;
}

function hashPassword(password,salt) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt+"|"+password, Utilities.Charset.UTF_8);
  return bytes.map(b=>("0"+((b<0?b+256:b).toString(16))).slice(-2)).join("");
}
function ensureCredentialHeaders(sh) {
  if(sh.getLastColumn()<20) sh.insertColumnsAfter(sh.getLastColumn(),20-sh.getLastColumn());
  sh.getRange(1,19).setValue("Password Hash");
  sh.getRange(1,20).setValue("Password Salt");
}

function registerMember(body,elevare) {
  if (elevare) setupElevareRegistrationStatus();
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const registrationCode=clean(body.registrationCode || body.code).toUpperCase();
    const codeCheck=elevare ? verifyElevareCode(registrationCode,true) : validateRegistrationCode(registrationCode);
    if(!codeCheck.ok) return codeCheck;
    if(elevare){
      if(!/^\d{6}$/.test(clean(body.otp))) return {ok:false,error:"Enter the email verification code."};
    } else if(!emailWasVerified(body.email)) return {ok:false,error:"Email has not been verified."};

    const username=clean(body.username);
    const password=String(body.password || "");
    if(username.length<3) return {ok:false,error:"Username must contain at least 3 characters."};
    if(password.length<8) return {ok:false,error:"Password must contain at least 8 characters."};

    ensureMemberExtraHeaders();
    const sh=SS.getSheetByName("Members"), codes=SS.getSheetByName("Registration Codes");
    if(!sh) return {ok:false,error:'Sheet "Members" not found.'};
    ensureCredentialHeaders(sh);

    const existing=sh.getDataRange().getValues();
    const codeRows=codes.getDataRange().getValues();
    const codeRow=codeRows.findIndex((row,index)=>index>0 && clean(row[0]).toUpperCase()===registrationCode)+1;
    const elevareColumn=headerMap(codes).elevarestatus;
    const pendingIndex=existing.findIndex((row,index)=>index>0 &&
      clean(row[9]).toUpperCase()===registrationCode &&
      clean(row[14]).toLowerCase()==="elevare only");
    if(pendingIndex>0){
      const pending=existing[pendingIndex];
      const sameAccount=clean(pending[1]).toLowerCase()===username.toLowerCase() &&
        clean(pending[6]).toLowerCase()===clean(body.email).toLowerCase() &&
        hashPassword(password,clean(pending[19]))===clean(pending[18]);
      if(!sameAccount) return {ok:false,error:"Use the username, password, and email from your Elevare account."};
      const memberId=clean(pending[0]);
      if(elevare){
        const pendingDependents=Array.isArray(body.dependents) ? body.dependents : [];
        saveElevareDependents(memberId,clean(pending[20]),pendingDependents);
        SpreadsheetApp.flush();
        codes.getRange(codeRow,elevareColumn).setValue("Used");
      } else {
        const now=new Date();
        const dependents=SS.getSheetByName("Dependents");
        if(dependents){
          const rows=dependents.getDataRange().getValues();
          for(let i=1;i<rows.length;i++){
            if(clean(rows[i][1])===memberId && clean(rows[i][10])==="Elevare Only"){
              dependents.getRange(i+1,8,1,4).setValues([[addDays(now,30),addDays(now,120),addDays(now,180),"Active"]]);
            }
          }
        }
        sh.getRange(pendingIndex+1,12,1,4).setValues([[addDays(now,30),addDays(now,120),addDays(now,180),"Active"]]);
        codes.getRange(codeRow,4,1,3).setValues([[true,memberId,now]]);
      }
      return {ok:true,memberId};
    }
    for(let i=1;i<existing.length;i++){
      if(clean(existing[i][1]).toLowerCase()===username.toLowerCase()) return {ok:false,error:"Username is already registered. Please choose another username."};
      if(clean(existing[i][6]).toLowerCase()===clean(body.email).toLowerCase()) return {ok:false,error:"Email address is already registered."};
    }

    const deps=Array.isArray(body.dependents)?body.dependents.filter(d=>clean(d.name||d.fullName)):[];
    const marital=clean(body.maritalStatus).toLowerCase();
    const accepted=[];
    for(const d of deps){
      const check=validateDependentRule(marital,d,accepted);
      if(!check.ok) return check;
      accepted.push(d.relationship);
    }

    if(elevare){
      const otpCheck=verifyEmailOtp(body.email,body.otp);
      if(!otpCheck.ok) return otpCheck;
    }
    const now=new Date(), memberId="CL-"+Utilities.getUuid().slice(0,8).toUpperCase();
    const salt=Utilities.getUuid();
    const hash=hashPassword(password,salt);

    const memberRow=[
      memberId, username, clean(body.name || body.fullName), clean(body.address),
      clean(body.phone || body.contact), clean(body.phone2 || body.additionalNumber),
      clean(body.email).toLowerCase(), clean(body.birthdate), clean(body.gender),
      registrationCode, now, elevare ? "" : addDays(now,30), elevare ? "" : addDays(now,120), elevare ? "" : addDays(now,180),
      elevare ? "Elevare Only" : "Active", "", "", "", hash, salt, clean(body.maritalStatus)
    ];
    const writtenRow=writeDatabaseRow(sh,memberRow,1);
    audit("REGISTER_MEMBER",memberId,"Members row "+writtenRow+" created for "+username,"SUCCESS");

    if(elevare){
      saveElevareDependents(memberId,marital,deps);
      SpreadsheetApp.flush();
      codes.getRange(codeRow,elevareColumn).setValue("Used");
      return {ok:true,memberId};
    }
    const rows=codes.getDataRange().getValues();
    for(let i=1;i<rows.length;i++){
      if(clean(rows[i][0]).toUpperCase()===registrationCode){
        codes.getRange(i+1,4).setValue(true);
        codes.getRange(i+1,5).setValue(memberId);
        codes.getRange(i+1,6).setValue(now);
        audit("USE_REGISTRATION_CODE",memberId,registrationCode,"SUCCESS");
        break;
      }
    }

    deps.forEach(d=>{ addDependent(memberId,d); });

    return {ok:true,memberId:memberId};
  } catch(err) { return {ok:false,stage:"registerMember",error:String(err.message||err)}; }
  finally { try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); } }
}

function loginMember(username,password,elevare) {
  try {
    username=clean(username).toLowerCase();
    password=String(password || "");
    const sh=SS.getSheetByName("Members");
    if(!sh) return {ok:false,error:'Sheet "Members" not found.'};
    ensureCredentialHeaders(sh);
    const rows=sh.getDataRange().getValues();

    for(let i=1;i<rows.length;i++){
      if(clean(rows[i][1]).toLowerCase()===username){
        if(!["active", ...(elevare ? ["elevare only"] : [])].includes(clean(rows[i][14]).toLowerCase())) return {ok:false,error:"This member account is not active."};
        const storedHash=clean(rows[i][18]), salt=clean(rows[i][19]);
        if(!storedHash || !salt) return {ok:false,error:"This account does not yet have login credentials. Please contact Carelink administration."};
        if(hashPassword(password,salt)!==storedHash) return {ok:false,error:"Incorrect username or password."};

        const token=Utilities.getUuid()+Utilities.getUuid();
        const session={
          memberId:clean(rows[i][0]), username:clean(rows[i][1]), name:clean(rows[i][2]),
          email:clean(rows[i][6]), status:clean(rows[i][14])
        };
        CacheService.getScriptCache().put((elevare ? "elevare_session_" : "session_")+token,JSON.stringify(session),SESSION_SECONDS);
        return {ok:true,token:token,member:session};
      }
    }
    return {ok:false,error:"Incorrect username or password."};
  } catch(err) { return {ok:false,stage:"login",error:String(err.message||err)}; }
}
function getSession(token,elevare) {
  token=clean(token); if(!token) return {ok:false,error:"No session."};
  const raw=CacheService.getScriptCache().get((elevare ? "elevare_session_" : "session_")+token);
  if(!raw) return {ok:false,error:"Session expired. Please log in again."};
  return {ok:true,member:JSON.parse(raw)};
}
function logoutMember(token) {
  token=clean(token);
  if(token) CacheService.getScriptCache().remove("session_"+token);
  return {ok:true};
}

function addDependent(memberId,d,elevare) {
  try {
    const sh=SS.getSheetByName("Dependents");
    if(!sh) return {ok:false,error:'Sheet "Dependents" not found.'};
    const now=new Date(), id="CLD-"+Utilities.getUuid().slice(0,8).toUpperCase();
    const depRow=[id,memberId,clean(d.name||d.fullName),clean(d.relationship),clean(d.birthdate),clean(d.gender),now,elevare ? "" : addDays(now,30),elevare ? "" : addDays(now,120),elevare ? "" : addDays(now,180),elevare ? "Elevare Only" : "Active"];
    const writtenRow=writeDatabaseRow(sh,depRow,1);
    audit("ADD_DEPENDENT",memberId,"Dependent "+id+" written to row "+writtenRow,"SUCCESS");
    return {ok:true,dependentId:id};
  } catch(err) { return {ok:false,stage:"addDependent",error:String(err.message||err)}; }
}


function normalizeHeader(s){ return clean(s).toLowerCase().replace(/[^a-z0-9]/g,""); }
function sheetObjects(sh){
  if(!sh) return [];
  const values=sh.getDataRange().getValues();
  if(values.length<2) return [];
  const headers=values[0].map(normalizeHeader);
  return values.slice(1).filter(r=>r.some(v=>clean(v)!=="")).map(r=>{
    const o={}; headers.forEach((h,i)=>{ if(h) o[h]=r[i]; }); return o;
  });
}
function pick(o,names){
  for(const n of names){ const k=normalizeHeader(n); if(o[k]!==undefined && clean(o[k])!=="") return o[k]; }
  return "";
}
function isoDate(v){
  if(v instanceof Date && !isNaN(v)) return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd");
  const d=new Date(v); return isNaN(d) ? clean(v) : Utilities.formatDate(d,Session.getScriptTimeZone(),"yyyy-MM-dd");
}
function safeMoney(v){
  if(typeof v==="number") return v;
  const n=Number(String(v||"").replace(/[^0-9.-]/g,"")); return isNaN(n)?0:n;
}
function getMemberPortalData(token){
  try{
    const sess=getSession(token);
    if(!sess.ok) return sess;
    const memberId=sess.member.memberId;
    const msh=SS.getSheetByName("Members");
    const rows=msh.getDataRange().getValues();
    let r=null;
    for(let i=1;i<rows.length;i++) if(clean(rows[i][0])===memberId){r=rows[i];break;}
    if(!r) return {ok:false,error:"Member record was not found."};

    const member={
      memberId:clean(r[0]),username:clean(r[1]),name:clean(r[2]),address:clean(r[3]),
      contact:clean(r[4]),additionalNumber:clean(r[5]),email:clean(r[6]),
      birthdate:isoDate(r[7]),gender:clean(r[8]),registrationCode:clean(r[9]),
      registrationDate:isoDate(r[10]),accidentEligible:isoDate(r[11]),
      normalEligible:isoDate(r[12]),hospitalEligible:isoDate(r[13]),status:clean(r[14]),maritalStatus:clean(r[20]),statusReason:clean(r[21]),benefitsPausedAt:isoDate(r[22])
    };

    const deps=sheetObjects(SS.getSheetByName("Dependents"))
      .filter(o=>clean(pick(o,["Member ID","MemberID"]))===memberId)
      .map(o=>({
        dependentId:clean(pick(o,["Dependent ID","DependentID"])),
        name:clean(pick(o,["Full Name","Name"])),
        relationship:clean(pick(o,["Relationship"])),
        birthdate:isoDate(pick(o,["Birthdate","Birth Date","Date of Birth"])),
        gender:clean(pick(o,["Gender"])),
        dateAdded:isoDate(pick(o,["Date Added","Added Date"])),
        accidentEligible:isoDate(pick(o,["Accident Death Eligible Date"])),
        normalEligible:isoDate(pick(o,["Normal Death Eligible Date"])),
        hospitalEligible:isoDate(pick(o,["Hospitalization Eligible Date"])),
        status:clean(pick(o,["Status"]))
      }));

    const contributions=sheetObjects(SS.getSheetByName("Contributions"))
      .filter(o=>clean(pick(o,["Member ID","MemberID"]))===memberId)
      .map(o=>({
        date:isoDate(pick(o,["Date","Contribution Date","Payment Date"])),
        reference:clean(pick(o,["Reference","Reference No","Reference Number","Contribution ID","Transaction ID"])),
        type:clean(pick(o,["Type","Contribution Type","Payment Type"])) || "Contribution",
        amount:safeMoney(pick(o,["Amount","Contribution Amount","Payment Amount","Gross Amount"])),
        deduction:safeMoney(pick(o,["Deduction","Deduction Amount","Deductions"])),
        netAmount:(clean(pick(o,["Net Amount","Net Contribution"]))!=="" ? safeMoney(pick(o,["Net Amount","Net Contribution"])) : Math.max(0,safeMoney(pick(o,["Amount","Contribution Amount","Payment Amount","Gross Amount"]))-safeMoney(pick(o,["Deduction","Deduction Amount","Deductions"])))),
        deductionDetails:clean(pick(o,["Deduction Details","Deduction Description","Remarks"])),
        status:clean(pick(o,["Status","Payment Status"])) || "Recorded",
        receipt:clean(pick(o,["Receipt","Receipt URL","Receipt Link"]))
      }))
      .sort((a,b)=>String(b.date).localeCompare(String(a.date)));

    const total=contributions.reduce((s,c)=>s+c.amount,0);
    const year=String(new Date().getFullYear());
    const yearTotal=contributions.filter(c=>String(c.date).startsWith(year)).reduce((s,c)=>s+c.amount,0);
    return {ok:true,member,dependents:deps,contributions,
      summary:{total:total,count:contributions.length,thisYear:yearTotal,latest:contributions.length?contributions[0]:null}};
  }catch(err){ return {ok:false,stage:"getMemberPortalData",error:String(err.message||err)}; }
}


function ensureMemberExtraHeaders(){
  const sh=SS.getSheetByName("Members"); if(!sh) return;
  if(sh.getMaxColumns()<23) sh.insertColumnsAfter(sh.getMaxColumns(),23-sh.getMaxColumns());
  sh.getRange(1,21).setValue("Marital Status");
  sh.getRange(1,22).setValue("Status Reason");
  sh.getRange(1,23).setValue("Benefits Paused At");
}
function getMemberRowById(memberId){
  const sh=SS.getSheetByName("Members"), rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++) if(clean(rows[i][0])===clean(memberId)) return {sh,row:i+1,values:rows[i]};
  return null;
}
function relationClass(v){
  const r=clean(v).toLowerCase();
  if(["spouse","life partner","partner"].includes(r)) return "spouse";
  if(["child","son","daughter"].includes(r)) return "child";
  if(["parent","father","mother"].includes(r)) return "parent";
  if(["sibling","brother","sister"].includes(r)) return "sibling";
  return r;
}
function ageOnDate(birthdate, onDate){
  const b=new Date(birthdate), d=onDate||new Date();
  if(isNaN(b)) return null;
  let age=d.getFullYear()-b.getFullYear();
  const md=d.getMonth()-b.getMonth();
  if(md<0 || (md===0 && d.getDate()<b.getDate())) age--;
  return age;
}
function validateDependentRule(maritalStatus, dependent, existingRelationships){
  const marital=clean(maritalStatus).toLowerCase();
  const rel=relationClass(dependent.relationship);
  const existing=(existingRelationships||[]).map(relationClass);
  if(!["spouse","child","parent","sibling"].includes(rel))
    return {ok:false,error:"Please select Parent, Spouse/Life Partner, Child, or Sibling."};
  if(rel==="spouse" && existing.includes("spouse"))
    return {ok:false,error:"Only one spouse or life partner may be registered."};
  if(marital==="married" && (rel==="parent" || rel==="sibling"))
    return {ok:false,error:"Married members may only add a spouse/life partner and children below 18 years old."};
  if((rel==="child" || rel==="sibling")){
    const age=ageOnDate(dependent.birthdate,new Date());
    if(age===null) return {ok:false,error:"A valid birthdate is required."};
    if(age>=18) return {ok:false,error:(rel==="child"?"Children":"Siblings")+" must be below 18 years old when added."};
  }
  return {ok:true};
}
function addPortalDependent(body){
  try{
    const sess=getSession(body.token); if(!sess.ok) return sess;
    const memberId=sess.member.memberId, d=body.dependent||{}, rel=relationClass(d.relationship);
    if(!["spouse","child","parent","sibling"].includes(rel)) return {ok:false,error:"Please select an allowed dependent relationship."};
    ensureMemberExtraHeaders();
    const mr=getMemberRowById(memberId); if(!mr) return {ok:false,error:"Member record not found."};
    const marital=clean(mr.values[20]).toLowerCase();
    const existing=sheetObjects(SS.getSheetByName("Dependents")).filter(o=>clean(pick(o,["Member ID"]))===memberId);
    const classes=existing.map(o=>relationClass(pick(o,["Relationship"])));
    const check=validateDependentRule(marital,d,classes);
    if(!check.ok) return check;
    return addDependent(memberId,{name:d.name,relationship:d.relationship,birthdate:d.birthdate,gender:d.gender});
  }catch(err){return {ok:false,stage:"addPortalDependent",error:String(err.message||err)};}
}
function fmtPeso(n){ return "PHP "+Number(n||0).toFixed(2); }
function downloadContributionStatement(token){
  try{
    const data=getMemberPortalData(token); if(!data.ok) return data;
    const doc=DocumentApp.create("Carelink Statement - "+data.member.memberId), b=doc.getBody();
    b.appendParagraph("CARELINK COMMUNITY WELFARE SERVICES INC.").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    b.appendParagraph("Member Contribution Statement");
    b.appendParagraph("Member: "+data.member.name);
    b.appendParagraph("Membership ID: "+data.member.memberId);
    b.appendParagraph("Generated: "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"MMMM d, yyyy h:mm a"));
    b.appendParagraph("");
    const table=[["Date","Reference","Type","Contribution","Deduction","Net","Status"]];
    (data.contributions||[]).forEach(c=>table.push([c.date||"",c.reference||"",c.type||"",fmtPeso(c.amount),fmtPeso(c.deduction),fmtPeso(c.netAmount),c.status||""]));
    if(table.length===1)table.push(["","","No contribution records","","","",""]);
    b.appendTable(table); b.appendParagraph("");
    const gross=(data.contributions||[]).reduce((s,c)=>s+Number(c.amount||0),0);
    const ded=(data.contributions||[]).reduce((s,c)=>s+Number(c.deduction||0),0);
    const net=(data.contributions||[]).reduce((s,c)=>s+Number(c.netAmount||0),0);
    b.appendParagraph("Total Contributions: "+fmtPeso(gross));
    b.appendParagraph("Total Deductions: "+fmtPeso(ded));
    b.appendParagraph("Net Contributions: "+fmtPeso(net));
    b.appendParagraph("This statement is generated from Carelink's recorded contribution data.");
    doc.saveAndClose();
    const f=DriveApp.getFileById(doc.getId()), pdf=f.getAs(MimeType.PDF).setName("Carelink_Statement_"+data.member.memberId+".pdf");
    f.setTrashed(true); audit("DOWNLOAD_STATEMENT",data.member.memberId,"Contribution statement generated","SUCCESS");
    return {ok:true,fileName:pdf.getName(),mimeType:"application/pdf",base64:Utilities.base64Encode(pdf.getBytes())};
  }catch(err){return {ok:false,stage:"downloadStatement",error:String(err.message||err)};}
}
function eligibleFor(subject,type){
  const now=new Date();now.setHours(0,0,0,0);
  const k=type==="Accident Death"?"accidentEligible":type==="Normal Death"?"normalEligible":"hospitalEligible";
  const d=new Date(subject[k]);if(isNaN(d))return false;d.setHours(0,0,0,0);return now>=d;
}
function requestAssistance(body){
  try{
    const sess=getSession(body.token);if(!sess.ok)return sess;
    const data=getMemberPortalData(body.token);if(!data.ok)return data;
    if(clean(data.member.status).toLowerCase()!=="active")return {ok:false,error:"Your membership is inactive. Benefit countdowns are paused until your contribution balance is restored."};
    const type=clean(body.assistanceType);
    if(!["Hospitalization","Accident Death","Normal Death"].includes(type))return {ok:false,error:"Invalid assistance type."};
    let subject=clean(body.subjectType)==="member"?data.member:(data.dependents||[]).find(d=>clean(d.dependentId)===clean(body.subjectId));
    if(!subject)return {ok:false,error:"Selected member/dependent not found."};
    if(!eligibleFor(subject,type))return {ok:false,error:"The selected person is not yet eligible for this assistance type."};
    const files=Array.isArray(body.attachments)?body.attachments:[];
    if(!files.length)return {ok:false,error:"Please attach at least one supporting document."};
    if(files.length>5)return {ok:false,error:"Maximum 5 attachments."};
    const blobs=[];let total=0;
    for(const f of files){const bytes=Utilities.base64Decode(f.base64||"");total+=bytes.length;if(bytes.length>5*1024*1024)return {ok:false,error:"Each attachment must be 5 MB or smaller."};blobs.push(Utilities.newBlob(bytes,f.mimeType||"application/octet-stream",clean(f.name)||"attachment"));}
    if(total>15*1024*1024)return {ok:false,error:"Total attachments must be 15 MB or smaller."};
    const rid="CAR-"+Utilities.getUuid().slice(0,8).toUpperCase(), sname=subject.name||data.member.name, sid=subject.dependentId||data.member.memberId;
    const text=["New Carelink Assistance Request","Request ID: "+rid,"Requested by: "+data.member.name+" ("+data.member.memberId+")","Assistance type: "+type,"For: "+sname+" ("+sid+")","Contact: "+data.member.contact,"Email: "+data.member.email,"Details: "+clean(body.details),"Submitted: "+new Date()].join("\n");
    MailApp.sendEmail({to:"carelinkcwsi@gmail.com",subject:"Carelink Assistance Request "+rid+" - "+type,body:text,attachments:blobs});
    let sh=SS.getSheetByName("Assistance Requests");if(!sh){sh=SS.insertSheet("Assistance Requests");sh.appendRow(["Request ID","Date","Member ID","Member Name","Assistance Type","Subject ID","Subject Name","Details","Status"]);}
    writeDatabaseRow(sh,[rid,new Date(),data.member.memberId,data.member.name,type,sid,sname,clean(body.details),"Submitted"],1);
    audit("REQUEST_ASSISTANCE",data.member.memberId,rid+" "+type+" for "+sid,"SUCCESS");
    return {ok:true,requestId:rid};
  }catch(err){return {ok:false,stage:"requestAssistance",error:String(err.message||err)};}
}

function authorizePdfStatement() {
  const doc=DocumentApp.create("Carelink Authorization Test");
  const file=DriveApp.getFileById(doc.getId());
  doc.getBody().appendParagraph("Carelink PDF statement authorization completed.");
  doc.saveAndClose();
  file.setTrashed(true);
  return "PDF statement permissions authorized.";
}

function authorizeEmail() {
  MailApp.sendEmail("carelinkcwsi@gmail.com","Carelink Email Authorization Test","Success! Carelink Apps Script is authorized to send email.");
}


// ============================================================
// v9.1 CONTRIBUTION BALANCE + BENEFIT PAUSE ENGINE
// ============================================================
function memberContributionBalance(memberId){
  return sheetObjects(SS.getSheetByName("Contributions"))
    .filter(o=>clean(pick(o,["Member ID","MemberID"]))===clean(memberId))
    .reduce((sum,o)=>{
      const netRaw=pick(o,["Net Amount","Net Contribution"]);
      const net=clean(netRaw)!=="" ? safeMoney(netRaw) : safeMoney(pick(o,["Amount","Contribution Amount"]))-safeMoney(pick(o,["Deduction","Deduction Amount"]));
      return sum+net;
    },0);
}
function setMemberBalanceInactive(memberId,when){
  ensureMemberExtraHeaders();
  const mr=getMemberRowById(memberId); if(!mr)return;
  const hm=headerMap(mr.sh), now=when||new Date();
  mr.sh.getRange(mr.row,hm[normalizeHeader("Status")]).setValue("Inactive");
  mr.sh.getRange(mr.row,hm[normalizeHeader("Status Reason")]).setValue("Balance Depleted");
  const pc=hm[normalizeHeader("Benefits Paused At")];
  if(pc && !clean(mr.sh.getRange(mr.row,pc).getValue())) mr.sh.getRange(mr.row,pc).setValue(now);
}
function shiftFutureEligibilityDate(value,pausedAt,resumedAt){
  const d=value instanceof Date?new Date(value):new Date(value);
  if(isNaN(d))return value;
  const p=new Date(pausedAt), r=new Date(resumedAt);
  if(isNaN(p)||isNaN(r)||d<=p)return value; // already eligible before pause stays eligible
  return new Date(d.getTime()+Math.max(0,r.getTime()-p.getTime()));
}
function reactivateMemberFromBalance(memberId,when){
  ensureMemberExtraHeaders();
  const mr=getMemberRowById(memberId); if(!mr)return false;
  const hm=headerMap(mr.sh), reason=clean(mr.sh.getRange(mr.row,hm[normalizeHeader("Status Reason")]).getValue());
  const status=clean(mr.sh.getRange(mr.row,hm[normalizeHeader("Status")]).getValue());
  if(status.toLowerCase()!=="inactive" || reason!=="Balance Depleted")return false;
  const pc=hm[normalizeHeader("Benefits Paused At")], paused=pc?mr.sh.getRange(mr.row,pc).getValue():"";
  const resumed=when||new Date();
  if(paused){
    ["Accident Death Eligible Date","Normal Death Eligible Date","Hospitalization Eligible Date"].forEach(h=>{
      const c=hm[normalizeHeader(h)]; if(c){const old=mr.sh.getRange(mr.row,c).getValue();mr.sh.getRange(mr.row,c).setValue(shiftFutureEligibilityDate(old,paused,resumed));}
    });
    const dsh=SS.getSheetByName("Dependents");
    if(dsh){
      const dh=headerMap(dsh), mc=dh[normalizeHeader("Member ID")];
      if(mc && dsh.getLastRow()>=2){
        const ids=dsh.getRange(2,mc,dsh.getLastRow()-1,1).getDisplayValues();
        ids.forEach((v,i)=>{if(clean(v[0])===clean(memberId)){
          ["Accident Death Eligible Date","Normal Death Eligible Date","Hospitalization Eligible Date"].forEach(h=>{
            const c=dh[normalizeHeader(h)]; if(c){const cell=dsh.getRange(i+2,c), old=cell.getValue();cell.setValue(shiftFutureEligibilityDate(old,paused,resumed));}
          });
        }});
      }
    }
  }
  mr.sh.getRange(mr.row,hm[normalizeHeader("Status")]).setValue("Active");
  mr.sh.getRange(mr.row,hm[normalizeHeader("Status Reason")]).setValue("");
  if(pc)mr.sh.getRange(mr.row,pc).setValue("");
  audit("AUTO_REACTIVATE",memberId,"Contribution balance restored above zero; benefit countdown resumed.","SUCCESS");
  return true;
}
function assessActiveMembersForAssistance(requestId,admin){
  const rid=clean(requestId); if(!rid)throw new Error("Request ID is required for assessment.");
  const ash=ensureSheetWithHeaders("Assistance Assessments",["Assessment ID","Request ID","Date","Member ID","Balance Before","Assessment","Balance After","Status","Recorded By"]);
  if(sheetObjects(ash).some(o=>clean(pick(o,["Request ID"]))===rid))return {alreadyAssessed:true,assessed:0,total:0};
  const csh=ensureSheetWithHeaders("Contributions",["Contribution ID","Member ID","Date","Reference","Type","Amount","Deduction","Net Amount","Deduction Details","Status"]);
  ensureMemberExtraHeaders();
  const members=sheetObjects(SS.getSheetByName("Members")).filter(o=>clean(pick(o,["Member ID"]))!=="");
  let count=0,total=0;
  members.forEach(o=>{
    const mid=clean(pick(o,["Member ID"])), status=clean(pick(o,["Status"])).toLowerCase();
    const before=Math.max(0,memberContributionBalance(mid));
    if(status!=="active" || before<=0)return; // only active, funded members at approval time
    const charge=Math.min(4,before), after=Math.max(0,before-charge), now=new Date();
    const cid="ASM-"+Utilities.getUuid().slice(0,8).toUpperCase();
    appendObject(csh,{"Contribution ID":cid,"Member ID":mid,"Date":now,"Reference":rid,"Type":"Assistance Assessment","Amount":0,"Deduction":charge,"Net Amount":-charge,"Deduction Details":"₱4 active-member assessment for approved assistance "+rid,"Status":"Recorded"},"Contribution ID");
    appendObject(ash,{"Assessment ID":"AA-"+Utilities.getUuid().slice(0,8).toUpperCase(),"Request ID":rid,"Date":now,"Member ID":mid,"Balance Before":before,"Assessment":charge,"Balance After":after,"Status":"Posted","Recorded By":admin.username},"Assessment ID");
    if(after<=0)setMemberBalanceInactive(mid,now);
    count++; total+=charge;
  });
  return {alreadyAssessed:false,assessed:count,total:total};
}

// ============================================================
// CARELINK ADMIN PORTAL v9.1
// ============================================================
const ADMIN_SESSION_SECONDS = 14400; // 4 hours

function ensureSheetWithHeaders(name, headers){
  let sh=SS.getSheetByName(name);
  if(!sh) sh=SS.insertSheet(name);
  if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  const current=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getDisplayValues()[0].map(clean);
  headers.forEach(h=>{ if(!current.some(x=>normalizeHeader(x)===normalizeHeader(h))){ sh.getRange(1,sh.getLastColumn()+1).setValue(h); current.push(h); } });
  return sh;
}
function headerMap(sh){
  const hs=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0]; const m={};
  hs.forEach((h,i)=>m[normalizeHeader(h)]=i+1); return m;
}
function appendObject(sh,obj,keyHeader){
  const hm=headerMap(sh), vals=new Array(sh.getLastColumn()).fill("");
  Object.keys(obj).forEach(k=>{const c=hm[normalizeHeader(k)];if(c)vals[c-1]=obj[k];});
  return writeDatabaseRow(sh,vals,hm[normalizeHeader(keyHeader||Object.keys(obj)[0])]||1);
}
function updateObjectByKey(sh,keyHeader,keyValue,obj){
  const hm=headerMap(sh), kc=hm[normalizeHeader(keyHeader)]; if(!kc) throw new Error("Missing key column: "+keyHeader);
  const last=sh.getLastRow(); if(last<2) return false;
  const vals=sh.getRange(2,kc,last-1,1).getDisplayValues();
  for(let i=0;i<vals.length;i++) if(clean(vals[i][0])===clean(keyValue)){
    Object.keys(obj).forEach(k=>{const c=hm[normalizeHeader(k)];if(c)sh.getRange(i+2,c).setValue(obj[k]);}); return true;
  }
  return false;
}
function setupAdminSystem(){
  const au=ensureSheetWithHeaders("Admin Users",["Admin ID","Username","Full Name","Role","Email","Status","Password Hash","Password Salt","Created Date","Last Login"]);
  ensureSheetWithHeaders("Fund Transactions",["Transaction ID","Date","Type","Reference","Member ID","Description","Money In","Money Out","Status","Recorded By"]);
  ensureSheetWithHeaders("Assistance Requests",["Request ID","Date","Member ID","Member Name","Assistance Type","Subject ID","Subject Name","Details","Status","Amount Approved","Amount Released","Release Date","Release Reference","Admin Remarks","Processed By"]);
  ensureSheetWithHeaders("Assistance Assessments",["Assessment ID","Request ID","Date","Member ID","Balance Before","Assessment","Balance After","Status","Recorded By"]);
  ensureMemberExtraHeaders();
  const users=sheetObjects(au);
  if(!users.length){
    const temp="CL-"+Utilities.getUuid().replace(/-/g,"").slice(0,10);
    const salt=Utilities.getUuid();
    appendObject(au,{"Admin ID":"ADM-001","Username":"admin","Full Name":"CareLink Administrator","Role":"Super Admin","Email":"carelinkcwsi@gmail.com","Status":"Active","Password Hash":hashPassword(temp,salt),"Password Salt":salt,"Created Date":new Date()},"Admin ID");
    Logger.log("CARELINK FIRST ADMIN -> username: admin | temporary password: "+temp);
  }
  return "Admin system ready. If this is the first setup, open Execution log to see the temporary admin password.";
}
function loginAdmin(username,password){
  const sh=ensureSheetWithHeaders("Admin Users",["Admin ID","Username","Full Name","Role","Email","Status","Password Hash","Password Salt","Created Date","Last Login"]);
  const hm=headerMap(sh), rows=sh.getDataRange().getValues(); username=clean(username).toLowerCase();
  for(let i=1;i<rows.length;i++){
    if(clean(rows[i][hm.username-1]).toLowerCase()===username){
      if(clean(rows[i][hm.status-1]).toLowerCase()!=="active") return {ok:false,error:"Admin account is not active."};
      const salt=clean(rows[i][hm.passwordsalt-1]), hash=clean(rows[i][hm.passwordhash-1]);
      if(hashPassword(String(password||""),salt)!==hash) return {ok:false,error:"Incorrect admin username or password."};
      const token=Utilities.getUuid()+Utilities.getUuid();
      const admin={adminId:clean(rows[i][hm.adminid-1]),username:clean(rows[i][hm.username-1]),name:clean(rows[i][hm.fullname-1]),role:clean(rows[i][hm.role-1]),email:clean(rows[i][hm.email-1])};
      CacheService.getScriptCache().put("admin_session_"+token,JSON.stringify(admin),ADMIN_SESSION_SECONDS);
      sh.getRange(i+1,hm.lastlogin).setValue(new Date());
      return {ok:true,token,admin};
    }
  }
  return {ok:false,error:"Incorrect admin username or password."};
}
function getAdminSession(token){
  const raw=CacheService.getScriptCache().get("admin_session_"+clean(token));
  if(!raw)return {ok:false,error:"Admin session expired. Please log in again."};
  return {ok:true,admin:JSON.parse(raw)};
}
function requireAdmin(token){const s=getAdminSession(token);if(!s.ok)throw new Error(s.error);return s.admin;}
function adminLogout(token){CacheService.getScriptCache().remove("admin_session_"+clean(token));return {ok:true};}
function adminAudit(admin,action,detail,status){audit("ADMIN_"+action,admin.adminId,(admin.username||"")+": "+detail,status||"SUCCESS");}

function adminDashboard(body){
  const admin=requireAdmin(body.token);
  const members=sheetObjects(SS.getSheetByName("Members")).filter(o=>clean(pick(o,["Member ID"]))!=="");
  const contributions=sheetObjects(SS.getSheetByName("Contributions"));
  const requests=sheetObjects(SS.getSheetByName("Assistance Requests"));
  const funds=sheetObjects(SS.getSheetByName("Fund Transactions"));
  const contributionTotal=contributions.filter(o=>clean(pick(o,["Type"])).toLowerCase()!=="assistance assessment").reduce((s,o)=>s+Math.max(0,safeMoney(pick(o,["Amount","Contribution Amount"]))),0);
  const manualIn=funds.reduce((s,o)=>s+safeMoney(pick(o,["Money In"])),0);
  const moneyOut=funds.reduce((s,o)=>s+safeMoney(pick(o,["Money Out"])),0);
  const balance=contributionTotal+manualIn-moneyOut;
  const pending=requests.filter(o=>!["released","declined","cancelled"].includes(clean(pick(o,["Status"])).toLowerCase())).length;
  return {ok:true,admin,summary:{availableFund:balance,totalContributions:contributionTotal,totalMoneyOut:moneyOut,pendingAssistance:pending,activeMembers:members.filter(o=>clean(pick(o,["Status"])).toLowerCase()==="active").length,totalMembers:members.length},recentFunds:funds.slice(-10).reverse().map(o=>({id:clean(pick(o,["Transaction ID"])),date:isoDate(pick(o,["Date"])),type:clean(pick(o,["Type"])),reference:clean(pick(o,["Reference"])),description:clean(pick(o,["Description"])),moneyIn:safeMoney(pick(o,["Money In"])),moneyOut:safeMoney(pick(o,["Money Out"]))}))};
}
function adminListMembers(body){
  requireAdmin(body.token); const q=clean(body.query).toLowerCase();
  return {ok:true,members:sheetObjects(SS.getSheetByName("Members")).map(o=>({memberId:clean(pick(o,["Member ID"])),username:clean(pick(o,["Username"])),name:clean(pick(o,["Full Name","Name"])),address:clean(pick(o,["Address"])),contact:clean(pick(o,["Contact","Phone"])),email:clean(pick(o,["Email"])),birthdate:isoDate(pick(o,["Birthdate"])),gender:clean(pick(o,["Gender"])),status:clean(pick(o,["Status"])),maritalStatus:clean(pick(o,["Marital Status"]))})).filter(x=>x.memberId && (!q||JSON.stringify(x).toLowerCase().includes(q))).slice(0,200)};
}
function adminUpdateMember(body){
  const admin=requireAdmin(body.token), mr=getMemberRowById(body.memberId); if(!mr)return {ok:false,error:"Member not found."};
  ensureMemberExtraHeaders(); const sh=mr.sh, row=mr.row;
  const allowed={name:3,address:4,contact:5,additionalNumber:6,email:7,birthdate:8,gender:9,status:15,maritalStatus:21};
  Object.keys(allowed).forEach(k=>{if(body.member&&body.member[k]!==undefined)sh.getRange(row,allowed[k]).setValue(clean(body.member[k]));});
  adminAudit(admin,"UPDATE_MEMBER",clean(body.memberId),"SUCCESS"); return {ok:true};
}
function adminListCodes(body){
  requireAdmin(body.token); const sh=SS.getSheetByName("Registration Codes");
  return {ok:true,codes:sheetObjects(sh).map(o=>({code:clean(pick(o,["Code"])),issuedDate:isoDate(pick(o,["Issued Date"])),status:clean(pick(o,["Status"])),used:boolVal(pick(o,["Used"])),usedBy:clean(pick(o,["Used By"])),usedDate:isoDate(pick(o,["Used Date"])),registrationPaid:boolVal(pick(o,["Registration Paid"])),initialDonationPaid:boolVal(pick(o,["Initial Donation Paid"])),notes:clean(pick(o,["Admin Notes"]))})).reverse().slice(0,200)};
}
function adminSaveCode(body){
  const admin=requireAdmin(body.token); const sh=ensureSheetWithHeaders("Registration Codes",["Code","Issued Date","Status","Used","Used By","Used Date","Registration Paid","Initial Donation Paid","Admin Notes"]);
  let code=clean(body.code).toUpperCase(); if(!code)code="CL-"+Math.random().toString(36).slice(2,8).toUpperCase();
  const found=updateObjectByKey(sh,"Code",code,{"Status":clean(body.status)||"Active","Registration Paid":!!body.registrationPaid,"Initial Donation Paid":!!body.initialDonationPaid,"Admin Notes":clean(body.notes)});
  if(!found)appendObject(sh,{"Code":code,"Issued Date":new Date(),"Status":clean(body.status)||"Active","Used":false,"Registration Paid":!!body.registrationPaid,"Initial Donation Paid":!!body.initialDonationPaid,"Admin Notes":clean(body.notes)},"Code");
  adminAudit(admin,"SAVE_CODE",code,"SUCCESS"); return {ok:true,code};
}
function adminListContributions(body){
  requireAdmin(body.token); return {ok:true,contributions:sheetObjects(SS.getSheetByName("Contributions")).map(o=>({id:clean(pick(o,["Contribution ID","Transaction ID","Reference"])),memberId:clean(pick(o,["Member ID"])),date:isoDate(pick(o,["Date","Contribution Date"])),reference:clean(pick(o,["Reference"])),type:clean(pick(o,["Type"])),amount:safeMoney(pick(o,["Amount"])),deduction:safeMoney(pick(o,["Deduction"])),netAmount:safeMoney(pick(o,["Net Amount"])),details:clean(pick(o,["Deduction Details"])),status:clean(pick(o,["Status"]))})).reverse().slice(0,300)};
}
function adminSaveContribution(body){
  const admin=requireAdmin(body.token), c=body.contribution||{}; if(!getMemberRowById(c.memberId))return {ok:false,error:"Member ID not found."};
  const sh=ensureSheetWithHeaders("Contributions",["Contribution ID","Member ID","Date","Reference","Type","Amount","Deduction","Net Amount","Deduction Details","Status"]);
  const id=clean(c.id)||"CON-"+Utilities.getUuid().slice(0,8).toUpperCase(), amount=safeMoney(c.amount), deduction=safeMoney(c.deduction), net=amount-deduction;
  if(amount<0||deduction<0)return {ok:false,error:"Amount and deduction cannot be negative."};
  const obj={"Contribution ID":id,"Member ID":clean(c.memberId),"Date":c.date?new Date(c.date+"T00:00:00"):new Date(),"Reference":clean(c.reference),"Type":clean(c.type)||"Contribution","Amount":amount,"Deduction":deduction,"Net Amount":net,"Deduction Details":clean(c.details),"Status":clean(c.status)||"Recorded"};
  if(!updateObjectByKey(sh,"Contribution ID",id,obj))appendObject(sh,obj,"Contribution ID");
  const balance=memberContributionBalance(c.memberId);
  if(balance>0)reactivateMemberFromBalance(c.memberId,new Date());
  else setMemberBalanceInactive(c.memberId,new Date());
  adminAudit(admin,"SAVE_CONTRIBUTION",id+" for "+clean(c.memberId),"SUCCESS"); return {ok:true,id,netAmount:net,balance:balance};
}
function adminListAssistance(body){
  requireAdmin(body.token); const sh=ensureSheetWithHeaders("Assistance Requests",["Request ID","Date","Member ID","Member Name","Assistance Type","Subject ID","Subject Name","Details","Status","Amount Approved","Amount Released","Release Date","Release Reference","Admin Remarks","Processed By"]);
  return {ok:true,requests:sheetObjects(sh).map(o=>({requestId:clean(pick(o,["Request ID"])),date:isoDate(pick(o,["Date"])),memberId:clean(pick(o,["Member ID"])),memberName:clean(pick(o,["Member Name"])),type:clean(pick(o,["Assistance Type"])),subjectId:clean(pick(o,["Subject ID"])),subjectName:clean(pick(o,["Subject Name"])),details:clean(pick(o,["Details"])),status:clean(pick(o,["Status"])),amountApproved:safeMoney(pick(o,["Amount Approved"])),amountReleased:safeMoney(pick(o,["Amount Released"])),releaseDate:isoDate(pick(o,["Release Date"])),releaseReference:clean(pick(o,["Release Reference"])),remarks:clean(pick(o,["Admin Remarks"]))})).reverse().slice(0,300)};
}
function adminUpdateAssistance(body){
  const admin=requireAdmin(body.token), a=body.assistance||{}, sh=ensureSheetWithHeaders("Assistance Requests",["Request ID","Date","Member ID","Member Name","Assistance Type","Subject ID","Subject Name","Details","Status","Amount Approved","Amount Released","Release Date","Release Reference","Admin Remarks","Processed By"]);
  const rid=clean(a.requestId); if(!rid)return {ok:false,error:"Request ID is required."};
  const current=sheetObjects(sh).find(o=>clean(pick(o,["Request ID"]))===rid); if(!current)return {ok:false,error:"Assistance request not found."};
  const oldStatus=clean(pick(current,["Status"])).toLowerCase(), status=clean(a.status)||"Under Review", newStatus=status.toLowerCase(), released=safeMoney(a.amountReleased);
  const ok=updateObjectByKey(sh,"Request ID",rid,{"Status":status,"Amount Approved":safeMoney(a.amountApproved),"Amount Released":released,"Release Date":newStatus==="released"?new Date():clean(a.releaseDate),"Release Reference":clean(a.releaseReference),"Admin Remarks":clean(a.remarks),"Processed By":admin.username});
  if(!ok)return {ok:false,error:"Assistance request not found."};
  let assessment=null;
  if(newStatus==="approved" && oldStatus!=="approved") assessment=assessActiveMembersForAssistance(rid,admin);
  if(newStatus==="released" && released>0){
    const fsh=ensureSheetWithHeaders("Fund Transactions",["Transaction ID","Date","Type","Reference","Member ID","Description","Money In","Money Out","Status","Recorded By"]);
    const existing=sheetObjects(fsh).some(o=>clean(pick(o,["Reference"]))===rid && clean(pick(o,["Type"])).toLowerCase()==="assistance released");
    if(!existing)appendObject(fsh,{"Transaction ID":"FT-"+Utilities.getUuid().slice(0,8).toUpperCase(),"Date":new Date(),"Type":"Assistance Released","Reference":rid,"Member ID":clean(a.memberId)||clean(pick(current,["Member ID"])),"Description":clean(a.type||pick(current,["Assistance Type"]))+" - "+clean(a.subjectName||pick(current,["Subject Name"])),"Money In":0,"Money Out":released,"Status":"Posted","Recorded By":admin.username},"Transaction ID");
  }
  adminAudit(admin,"UPDATE_ASSISTANCE",rid+" -> "+status,"SUCCESS"); return {ok:true,assessment:assessment};
}
function adminAddFundTransaction(body){
  const admin=requireAdmin(body.token), x=body.transaction||{}; const mi=safeMoney(x.moneyIn), mo=safeMoney(x.moneyOut); if(mi<0||mo<0||(!mi&&!mo))return {ok:false,error:"Enter a Money In or Money Out amount."};
  const sh=ensureSheetWithHeaders("Fund Transactions",["Transaction ID","Date","Type","Reference","Member ID","Description","Money In","Money Out","Status","Recorded By"]), id="FT-"+Utilities.getUuid().slice(0,8).toUpperCase();
  appendObject(sh,{"Transaction ID":id,"Date":x.date?new Date(x.date+"T00:00:00"):new Date(),"Type":clean(x.type)||"Adjustment","Reference":clean(x.reference),"Member ID":clean(x.memberId),"Description":clean(x.description),"Money In":mi,"Money Out":mo,"Status":"Posted","Recorded By":admin.username},"Transaction ID");
  adminAudit(admin,"FUND_TRANSACTION",id,"SUCCESS"); return {ok:true,id};
}

/** Run once in the Apps Script editor, or let the first code check create the column. */
function setupElevareRegistrationStatus() {
  const sheet = SS.getSheetByName("Registration Codes");
  if (!sheet) throw new Error('Sheet "Registration Codes" not found.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = headerMap(sheet).elevarestatus;
    if (existing) return existing;
    const column = sheet.getLastColumn() + 1;
    if (column > sheet.getMaxColumns()) sheet.insertColumnAfter(column - 1);
    sheet.getRange(1, column).setValue("Elevare Status");
    SpreadsheetApp.flush();
    return column;
  } finally {
    lock.releaseLock();
  }
}

function verifyElevareCode(value,registrationLockHeld) {
  const code = clean(value).toUpperCase();
  const denied = {ok:false,error:"This code is unavailable for Elevare registration."};
  if (!/^[A-Z0-9-]{8,100}$/.test(code)) return denied;

  const column = registrationLockHeld ? headerMap(SS.getSheetByName("Registration Codes")).elevarestatus : setupElevareRegistrationStatus();
  const sheet = SS.getSheetByName("Registration Codes");
  const rows = sheet.getDataRange().getValues().slice(1);
  const matches = rows.filter(row => clean(row[0]).toUpperCase() === code);
  if (matches.length !== 1) return denied;

  const row = matches[0];
  if (clean(row[2]).toLowerCase() !== "active" ||
      clean(row[column - 1]).toLowerCase() !== "paid") return denied;
  return {ok:true};
}



function saveElevareDependents(memberId, marital, dependents) {
  if (!dependents.length) return;
  const sheet = SS.getSheetByName("Dependents");
  if (!sheet) throw new Error('Sheet "Dependents" not found.');
  const existing = sheet.getDataRange().getValues().slice(1)
    .filter(row => clean(row[1]) === memberId);
  const relationships = existing.map(row => row[3]);
  for (const dependent of dependents) {
    const name = clean(dependent.name || dependent.fullName);
    if (!name) continue;
    const birthdate = clean(dependent.birthdate);
    const same = existing.some(row => clean(row[2]) === name &&
      relationClass(row[3]) === relationClass(dependent.relationship) &&
      isoDate(row[4]) === isoDate(birthdate));
    if (same) continue;
    const check = validateDependentRule(marital, dependent, relationships);
    if (!check.ok) throw new Error(check.error);
    const saved = addDependent(memberId, dependent, true);
    if (!saved.ok) throw new Error(saved.error || "Unable to save dependent.");
    relationships.push(dependent.relationship);
  }
}
