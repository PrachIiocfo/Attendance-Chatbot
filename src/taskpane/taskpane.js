/* taskpane.js — Attendance Assistant (real Jan-Apr 2026 data: A/P/WO, time-based late) */

const OLLAMA_MODEL = "gemma3:4b";
const EMBED_MODEL  = "nomic-embed-text";
let POLICY_CHUNKS = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("send").onclick = ask;
    document.getElementById("question").addEventListener("keydown", (e) => { if (e.key === "Enter") ask(); });
    document.querySelectorAll(".chip").forEach((chip) => {
      chip.onclick = () => { document.getElementById("question").value = chip.textContent; ask(); };
    });
  }
});

function fmtTime(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number" && v > 0 && v < 1) {
    let mins = Math.round(v * 24 * 60);
    return String(Math.floor(mins / 60)).padStart(2, "0") + ":" + String(mins % 60).padStart(2, "0");
  }
  return String(v).trim();
}
function fmtDate(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number" || /^\d{5}(\.\d+)?$/.test(String(v).trim())) {
    const d = new Date((Number(v) - 25569) * 86400 * 1000);
    if (!isNaN(d)) return String(d.getUTCDate()).padStart(2,"0")+"-"+String(d.getUTCMonth()+1).padStart(2,"0")+"-"+d.getUTCFullYear();
  }
  return String(v).trim();
}

async function readSheetData() {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
    range.load("values");
    await context.sync();
    const vals = range.values;
    if (!vals || vals.length < 2) return [];
    const header = vals[0].map((h) => String(h || "").trim().toLowerCase());
    const find = (...keys) => {
      for (let i = 0; i < header.length; i++)
        if (keys.some((k) => header[i] === k || header[i].includes(k))) return i;
      return -1;
    };
    const col = {
      date: find("date"),
      employee: find("employee", "name", "emp"),
      login: find("login", "in time", "intime"),
      logout: find("logout", "out time", "outtime"),
      hours: find("hours", "worked", "total"),
      status: find("status"),
      late: find("late"),
      early: find("early"),
    };
    const body = vals.slice(1);
    return body.map((r) => ({
      Date: fmtDate(col.date >= 0 ? r[col.date] : ""),
      Employee: String((col.employee >= 0 ? r[col.employee] : "") == null ? "" : (col.employee >= 0 ? r[col.employee] : "")).trim(),
      Login: fmtTime(col.login >= 0 ? r[col.login] : ""),
      Logout: fmtTime(col.logout >= 0 ? r[col.logout] : ""),
      Hours: String((col.hours >= 0 ? r[col.hours] : "") || "").trim(),
      Status: String((col.status >= 0 ? r[col.status] : "") || "").trim(),
      Late: fmtTime(col.late >= 0 ? r[col.late] : ""),
      Early: fmtTime(col.early >= 0 ? r[col.early] : ""),
    })).filter((r) => r.Employee && r.Employee.toLowerCase() !== "employee");
  });
}

async function loadPolicies() {
  try { const res = await fetch("assets/hr-policy.txt"); return res.ok ? await res.text() : ""; }
  catch (e) { return ""; }
}
function chunkText(t){return t.split(/\n\s*\n/).map(c=>c.trim()).filter(c=>c.length>30);}
async function embed(text){
  const res=await fetch("http://localhost:11434/api/embeddings",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:EMBED_MODEL,prompt:text})});
  const d=await res.json(); return d.embedding;
}
function cosine(a,b){let dot=0,na=0,nb=0;for(let i=0;i<a.length;i++){dot+=a[i]*b[i];na+=a[i]*a[i];nb+=b[i]*b[i];}return dot/(Math.sqrt(na)*Math.sqrt(nb)+1e-9);}
async function buildPolicyIndex(){
  if(POLICY_CHUNKS) return POLICY_CHUNKS;
  const t=await loadPolicies(); if(!t) return [];
  const out=[]; for(const c of chunkText(t)) out.push({text:c,vector:await embed(c)});
  POLICY_CHUNKS=out; return out;
}
async function retrieve(q,n){
  const idx=await buildPolicyIndex(); if(idx.length===0) return [];
  const qv=await embed(q);
  return idx.map(c=>({text:c.text,score:cosine(qv,c.vector)})).sort((a,b)=>b.score-a.score).slice(0,n||3);
}

function listEmployees(rows){return [...new Set(rows.map(r=>r.Employee))];}
function findName(rows,q){
  const ql=q.toLowerCase();
  const emps=listEmployees(rows);
  let m=ql.match(/(?:employee|emp|code)\s*#?\s*(\d+)/);
  if(m && emps.includes(m[1])) return m[1];
  const names=emps.filter(e=>/[a-z]/i.test(e)).sort((a,b)=>b.length-a.length);
  for(const e of names) if(ql.includes(e.toLowerCase())) return e;
  let cleaned=ql
    .replace(/\b\d{4}\b/g," ")
    .replace(/\b\d{1,2}\s*(st|nd|rd|th)?\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*/g," ")
    .replace(/(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s*\d{1,2}/g," ");
  const nums=cleaned.match(/\b\d+\b/g)||[];
  for(const n of nums) if(emps.includes(n)) return n;
  return null;
}
function monthInQ(q){
  const names={january:"01",february:"02",march:"03",april:"04",may:"05",june:"06",july:"07",august:"08",september:"09",october:"10",november:"11",december:"12",jan:"01",feb:"02",mar:"03",apr:"04",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};
  const ql=q.toLowerCase();
  for(const k in names){ if(new RegExp("\\b"+k+"\\b").test(ql)) return names[k]; }
  return null;
}
function dayInQ(q){
  const ql=q.toLowerCase();
  let m=ql.match(/\b(\d{1,2})\s*(st|nd|rd|th)?\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/);
  if(m) return String(Number(m[1]));
  m=ql.match(/(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s*(\d{1,2})/);
  if(m) return String(Number(m[2]));
  return null;
}
function dMonth(d){const m=String(d).match(/^(\d{2})-(\d{2})-/);return m?m[2]:"";}
function dDay(d){const m=String(d).match(/^(\d{2})-/);return m?String(Number(m[1])):"";}

function isPresent(s){return s==="P"||s==="WOP";}
function isAbsent(s){return s==="A";}
function isHalf(s){return s==="½P"||s==="WO½P";}
function isWO(s){return s==="WO"||s==="WO½P"||s==="WOP";}

function rankAll(rows,label,test){
  const c=listEmployees(rows).map(e=>({e,n:rows.filter(r=>r.Employee===e&&test(r)).length}));
  c.sort((a,b)=>b.n-a.n);
  const top=c.filter(x=>x.n>0).slice(0,8);
  return c[0].e+" had the most "+label+" ("+c[0].n+").\n\nTop:\n"+top.map(x=>x.e+": "+x.n).join("\n");
}

function attendancePct(rows, name, mon){
  const er=rows.filter(r=>r.Employee===name && (mon?dMonth(r.Date)===mon:true));
  if(er.length===0) return null;
  const present=er.filter(r=>isPresent(r.Status)).length;
  const half=er.filter(r=>isHalf(r.Status)).length;
  const pureWO=er.filter(r=>r.Status==="WO").length;
  const working=er.length - pureWO;
  const eff=present + 0.5*half;
  const pct=working? (100*eff/working) : 0;
  return {present, half, working, pct};
}

function compareTwo(rows, a, b, mon){
  const sa=attendancePct(rows,a,mon), sb=attendancePct(rows,b,mon);
  if(!sa||!sb) return "Could not find data for one of the employees.";
  const lateA=rows.filter(r=>r.Employee===a&&r.Late!==""&&(mon?dMonth(r.Date)===mon:true)).length;
  const lateB=rows.filter(r=>r.Employee===b&&r.Late!==""&&(mon?dMonth(r.Date)===mon:true)).length;
  return "Comparison"+(mon?" (selected month)":"")+":\n\n"+
    a+" \u2192 Present: "+sa.present+", Late: "+lateA+", Attendance: "+sa.pct.toFixed(1)+"%\n"+
    b+" \u2192 Present: "+sb.present+", Late: "+lateB+", Attendance: "+sb.pct.toFixed(1)+"%";
}

function rankByPct(rows, mon, ascending){
  const emps=listEmployees(rows);
  const arr=emps.map(e=>{ const st=attendancePct(rows,e,mon); return {e, pct: st? st.pct : 0}; });
  arr.sort((a,b)=> ascending ? a.pct-b.pct : b.pct-a.pct);
  return arr;
}

function teamStats(rows, mon){
  const er=rows.filter(r=>(mon?dMonth(r.Date)===mon:true));
  if(er.length===0) return null;
  const present=er.filter(r=>isPresent(r.Status)).length;
  const half=er.filter(r=>isHalf(r.Status)).length;
  const absent=er.filter(r=>isAbsent(r.Status)).length;
  const pureWO=er.filter(r=>r.Status==="WO").length;
  const late=er.filter(r=>r.Late!=="").length;
  const working=er.length - pureWO;
  const pct=working? (100*(present+0.5*half)/working) : 0;
  return {present, absent, late, working, pct, emps:listEmployees(er).length};
}

function monthRangeInQ(q){
  const map={january:"01",february:"02",march:"03",april:"04",may:"05",june:"06",july:"07",august:"08",september:"09",october:"10",november:"11",december:"12",jan:"01",feb:"02",mar:"03",apr:"04",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};
  const ql=q.toLowerCase();
  const m=ql.match(/(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)\s*(?:to|-|till|through|until|thru)\s*(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)/);
  if(!m) return null;
  let a=map[m[1]], b=map[m[2]];
  if(!a||!b) return null;
  if(a>b){ const t=a; a=b; b=t; }
  return {from:a, to:b};
}

function latePolicyCheck(rows, name, mon){
  const er=rows.filter(r=>r.Employee===name && r.Late!=="" && (mon?dMonth(r.Date)===mon:true));
  const perMonth={};
  er.forEach(r=>{ const m=dMonth(r.Date); perMonth[m]=(perMonth[m]||0)+1; });
  const months=Object.keys(perMonth).sort();
  const violatedMonths=months.filter(m=>perMonth[m]>3);
  const deduction=violatedMonths.length;
  let lines=months.map(m=>"Month "+m+": "+perMonth[m]+" late"+(perMonth[m]>3?"  \u26a0\ufe0f over limit":""));
  let out=name+(mon?" (selected month)":"")+" late-policy check:\n"+
    "Policy: more than 3 late arrivals in a month = 1 day leave deducted.\n\n"+
    (lines.length?lines.join("\n"):"No late arrivals.")+"\n\n";
  if(deduction>0) out+="\u26a0\ufe0f Policy exceeded in "+deduction+" month(s). Approx "+deduction+" day(s) of leave may be deducted.";
  else out+="\u2705 Within policy. No leave deduction.";
  return out;
}

// simple text bar for a percentage (0-100) -> filled blocks
function bar(pct){
  const full=Math.round((pct/100)*20);
  return "\u2588".repeat(Math.max(0,full)) + "\u2591".repeat(Math.max(0,20-full));
}
// only show bars when the user asks for a chart/graph/bar/visual
function wantsChart(ql){
  return ql.includes("chart")||ql.includes("graph")||ql.includes("bar")||ql.includes("visual")||ql.includes("plot");
}

function answerData(rows,q){
  const ql=q.toLowerCase();
  const name=findName(rows,q);
  const mon=monthInQ(q);
  const day=dayInQ(q);

  const range=monthRangeInQ(q);
  if(name && range){
    const inRange=(r)=>{const mm=dMonth(r.Date); return mm>=range.from && mm<=range.to;};
    const er=rows.filter(r=>r.Employee===name && inRange(r));
    if(er.length===0) return "No records found for "+name+" in that range.";
    const present=er.filter(r=>isPresent(r.Status)).length;
    const absent=er.filter(r=>isAbsent(r.Status)).length;
    const half=er.filter(r=>isHalf(r.Status)).length;
    const wo=er.filter(r=>isWO(r.Status)).length;
    const late=er.filter(r=>r.Late!=="").length;
    const early=er.filter(r=>r.Early!=="").length;
    const pureWO=er.filter(r=>r.Status==="WO").length;
    const working=er.length-pureWO;
    const pct=working?(100*(present+0.5*half)/working):0;
    return name+" (month "+range.from+" to "+range.to+") summary:\n"+
      "Present: "+present+"\nAbsent: "+absent+"\nHalf days: "+half+
      "\nWeekly offs: "+wo+"\nLate arrivals: "+late+"\nEarly departures: "+early+
      "\nAttendance: "+pct.toFixed(1)+"%\nTotal records: "+er.length;
  }

  // STATUS LIST
  if((ql.includes("list")||ql.includes("show all")||ql.includes("who all")||ql.includes("everyone who")||ql.includes("names of")||(ql.includes("who")&&(day||mon)))){
    const monTest2=(r)=>(mon?dMonth(r.Date)===mon:true);
    const dayTest2=(r)=>(day?dDay(r.Date)===day:true);
    const both2=(r)=>monTest2(r)&&dayTest2(r);
    let test=null, lab="";
    if(ql.includes("absent")||ql.includes("absence")){test=(r)=>isAbsent(r.Status)&&both2(r);lab="absent";}
    else if(ql.includes("late")){test=(r)=>r.Late!==""&&both2(r);lab="late";}
    else if(ql.includes("early")){test=(r)=>r.Early!==""&&both2(r);lab="left early";}
    else if(ql.includes("present")){test=(r)=>isPresent(r.Status)&&both2(r);lab="present";}
    else if(ql.includes("half")){test=(r)=>isHalf(r.Status)&&both2(r);lab="on half day";}
    if(test && (day||mon)){
      const who=[...new Set(rows.filter(test).map(r=>r.Employee))];
      const when=day?("on "+(day)+"/"+(mon||"")) : (mon?"in month "+mon:"");
      if(who.length===0) return "No one was "+lab+" "+when+".";
      return who.length+" employee(s) "+lab+" "+when+":\n\n"+who.join(", ");
    }
  }

  if(name && (ql.includes("violat")||ql.includes("policy")||ql.includes("deduct")||ql.includes("exceed")||ql.includes("break")||ql.includes("breach"))&&(ql.includes("late")||ql.includes("policy"))){
    return latePolicyCheck(rows, name, mon);
  }

  if(ql.includes("compare")){
    const emps=listEmployees(rows);
    const found=[];
    let cm=ql.match(/(?:employee|emp|code)\s*#?\s*(\d+)/g);
    if(cm) cm.forEach(x=>{const n=x.match(/\d+/)[0]; if(emps.includes(n)&&!found.includes(n)) found.push(n);});
    emps.filter(e=>/[a-z]/i.test(e)).forEach(e=>{ if(ql.includes(e.toLowerCase())&&!found.includes(e)) found.push(e); });
    if(found.length>=2) return compareTwo(rows, found[0], found[1], mon);
    return "To compare, mention two employees, e.g. 'Compare Employee 5 and Employee 20'.";
  }

  if((ql.includes("top")||ql.includes("bottom")||ql.includes("best")||ql.includes("worst")||ql.includes("least"))&&(ql.includes("attendance")||ql.includes("present")||ql.includes("punctual")||ql.includes("%"))){
    const ascending = ql.includes("bottom")||ql.includes("worst")||ql.includes("least");
    let nm=ql.match(/\b(\d{1,2})\b/); let topN=nm?Number(nm[1]):5; if(topN<1||topN>20) topN=5;
    const arr=rankByPct(rows,mon,ascending).slice(0,topN);
    const title=(ascending?"Bottom ":"Top ")+topN+" by attendance %"+(mon?" (selected month)":"")+":";
    if(wantsChart(ql))
      return title+"\n\n"+arr.map((x,i)=>(i+1)+". "+x.e+"\n"+bar(x.pct)+" "+x.pct.toFixed(1)+"%").join("\n\n");
    return title+"\n\n"+arr.map((x,i)=>(i+1)+". "+x.e+" \u2014 "+x.pct.toFixed(1)+"%").join("\n");
  }

  if((ql.includes("team")||ql.includes("overall")||ql.includes("everyone")||ql.includes("all employees")||ql.includes("company")||ql.includes("office"))&&(ql.includes("attendance")||ql.includes("present")||ql.includes("absent")||ql.includes("stat")||ql.includes("%")||ql.includes("late"))&&!ql.includes("timing")&&!ql.includes("hour")){
    const t=teamStats(rows,mon);
    if(!t) return "No records found"+(mon?" for the selected month":"")+".";
    return "Overall team"+(mon?" (selected month)":"")+":\n"+
      "Employees: "+t.emps+"\nTotal present days: "+t.present+"\nTotal absences: "+t.absent+
      "\nTotal late arrivals: "+t.late+"\nTeam attendance: "+t.pct.toFixed(1)+"%"+(wantsChart(ql)?("\n"+bar(t.pct)):"");
  }

  if(name && (ql.includes("percent")||ql.includes("%")||ql.includes("attendance rate")||ql.includes("attendance percentage"))){
    const st=attendancePct(rows,name,mon);
    if(!st) return "No records found for "+name+".";
    const barLine = wantsChart(ql) ? ("\n"+bar(st.pct)) : "";
    return name+(mon?" (selected month)":"")+" attendance: "+st.pct.toFixed(1)+"%"+barLine+"\n("+st.present+" present out of "+st.working+" working days).";
  }

  if(name&&mon&&day&&(ql.includes("login")||ql.includes("logout")||ql.includes("time")||ql.includes("status")||ql.includes("detail")||ql.includes("record")||ql.includes("hours")||ql.includes("late")||ql.includes("early"))){
    const row=rows.find(r=>r.Employee===name&&dMonth(r.Date)===mon&&dDay(r.Date)===day);
    if(!row) return "No record found for "+name+" on that day.";
    let out=name+" on "+row.Date+" — Status: "+(row.Status||"-");
    if(row.Login) out+=", Login: "+row.Login;
    if(row.Logout) out+=", Logout: "+row.Logout;
    if(row.Hours) out+=", Hours: "+row.Hours;
    if(row.Late) out+=", Late arrival at "+row.Late;
    if(row.Early) out+=", Early departure at "+row.Early;
    return out+".";
  }

  if(name && (ql.includes("record")||ql.includes("summary")||ql.includes("overview")||ql.includes("report"))){
    const mt=(r)=>(mon?dMonth(r.Date)===mon:true);
    const er=rows.filter(r=>r.Employee===name && mt(r));
    if(er.length===0) return "No records found for "+name+(mon?" in the selected month":"")+".";
    const present=er.filter(r=>isPresent(r.Status)).length;
    const absent=er.filter(r=>isAbsent(r.Status)).length;
    const half=er.filter(r=>isHalf(r.Status)).length;
    const wo=er.filter(r=>isWO(r.Status)).length;
    const late=er.filter(r=>r.Late!=="").length;
    const early=er.filter(r=>r.Early!=="").length;
    return name+(mon?" (selected month)":"")+" summary:\n"+
      "Present: "+present+"\nAbsent: "+absent+"\nHalf days: "+half+
      "\nWeekly offs: "+wo+"\nLate arrivals: "+late+"\nEarly departures: "+early+
      "\nTotal records: "+er.length;
  }

  const monTest=(r)=>(mon?dMonth(r.Date)===mon:true);
  const dayTest=(r)=>(day?dDay(r.Date)===day:true);
  const both=(r)=>monTest(r)&&dayTest(r);

  let metric=null,label="";
  if(ql.includes("late")){metric=(r)=>r.Late!==""&&both(r);label="late arrivals";}
  else if(ql.includes("early")){metric=(r)=>r.Early!==""&&both(r);label="early departures";}
  else if(ql.includes("absent")||ql.includes("absence")||ql.includes("leave")){metric=(r)=>isAbsent(r.Status)&&both(r);label="absences";}
  else if(ql.includes("half")){metric=(r)=>isHalf(r.Status)&&both(r);label="half days";}
  else if(ql.includes("weekly")||ql.includes("week off")||ql.includes("wo")){metric=(r)=>isWO(r.Status)&&both(r);label="weekly offs";}
  else if(ql.includes("present")){metric=(r)=>isPresent(r.Status)&&both(r);label="present days";}

  if(metric){
    if(ql.includes("who")||ql.includes("most")||ql.includes("highest")){
      const c=listEmployees(rows).map(e=>({e,n:rows.filter(r=>r.Employee===e&&metric(r)).length})).sort((a,b)=>b.n-a.n);
      if(ql.includes("who")&&day){
        const matched=c.filter(x=>x.n>0).map(x=>x.e);
        return matched.length?matched.join(", ")+" had "+label+" on that day.":"No one had "+label+" on that day.";
      }
      return rankAll(rows,label,metric);
    }
    if(name) return name+" had "+rows.filter(r=>r.Employee===name&&metric(r)).length+" "+label+(mon?" in the selected month":"")+".";
    return rankAll(rows,label,metric);
  }
  if(/(employee|emp|code)\s*#?\s*\d+/.test(ql) && !name){
    return "I couldn't find that employee in the sheet. Please check the code/name.";
  }
  return null;
}

async function answerDataWithAI(rows,q){
  const sample=rows.slice(0,60).map(r=>[r.Date,r.Employee,r.Login,r.Logout,r.Hours,r.Status,r.Late,r.Early].join(" | ")).join("\n");
  const prompt="You are an attendance assistant. Use ONLY this data. Columns: Date|Employee|Login|Logout|Hours|Status|Late|Early. Status: A=Absent, P=Present, WO=Weekly Off, \u00bdP=Half day. Late/Early columns show a time if the person was late/left early, else blank.\nDATA:\n"+sample+"\n\nQUESTION: "+q;
  return await callAI(prompt);
}

function routeQuestion(q, rows){
  const ql=q.toLowerCase();
  let namedPerson=/(employee|emp|code)\s*#?\s*\d+/.test(ql);
  if(!namedPerson && rows){
    namedPerson = listEmployees(rows).some(e=>/[a-z]/i.test(e) && ql.includes(e.toLowerCase()));
  }
  if((ql.includes("late")||ql.includes("violat")||ql.includes("deduct")||ql.includes("exceed")||ql.includes("breach")) && namedPerson) return "DATA";
  const policyWords=["policy","rule","entitle","eligible","allowed","able to","do we get","how to apply",
    "apply for","how do i","application process","wfh","work from home","casual leave","sick leave",
    "privilege","earned leave","office timing","office hour","working hour","weekly off policy","grace",
    "carry forward","notice period","comp off","zoho","what happens","what if","can i"];
  if(policyWords.some(w=>ql.includes(w))) return "POLICY";
  return "DATA";
}

async function ask(){
  const input=document.getElementById("question");
  const question=input.value.trim();
  if(!question) return;
  addMessage(question,"user"); input.value="";
  const thinking=addThinking(); setSending(true);
  try{
    const rows=await readSheetData();
    const route=routeQuestion(question, rows);
    let answer;
    if(route==="POLICY"){
      const top=await retrieve(question,3);
      if(top.length===0) answer="No policy file found (assets/hr-policy.txt).";
      else{
        const ctx=top.map((t,i)=>"["+(i+1)+"] "+t.text).join("\n\n");
        answer=await callAI("Answer concisely using ONLY the policy excerpts below. If not covered, say so.\n\nPOLICY:\n"+ctx+"\n\nQUESTION: "+question);
      }
    } else {
      answer=answerData(rows,question);
      if(answer===null) answer=await answerDataWithAI(rows,question);
    }
    thinking.remove(); addMessage(answer,"bot");
  }catch(err){ thinking.remove(); addMessage("Sorry, something went wrong: "+(err.message||err),"bot"); }
  finally{ setSending(false); }
}

async function callAI(prompt){
  try{
    const res=await fetch("http://localhost:11434/api/generate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:OLLAMA_MODEL,prompt:prompt,stream:false,options:{temperature:0.1}})});
    const d=await res.json(); return d.response?d.response.trim():"No answer returned.";
  }catch(e){ return "Could not reach the local AI. Make sure Ollama is running."; }
}

function addMessage(text,who){const chat=document.getElementById("chat");const div=document.createElement("div");div.className="msg "+who;div.textContent=text;chat.appendChild(div);chat.scrollTop=chat.scrollHeight;return div;}
function addThinking(){const chat=document.getElementById("chat");const div=document.createElement("div");div.className="thinking";div.textContent="thinking...";chat.appendChild(div);chat.scrollTop=chat.scrollHeight;return div;}
function setSending(on){document.getElementById("send").disabled=on;}