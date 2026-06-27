import React from "react";
import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";


var ICONS = ["🌅","☀️","🌞","🌆","🌇","🌃","🌙","⭐","🌟","✨","🔴","🟠","🟡","🟢","🔵","🟣","📋","🎯","🔆","💡"];
var DAYS  = [1,2,3,4,5,6,7,8,9,10];

var C = {
  navy:"#0F2D4A", amber:"#E67E22", bg:"#EEF2F7", card:"#FFF",
  text:"#1A202C", muted:"#718096", green:"#27AE60", red:"#E74C3C",
  purple:"#7C3AED", blue:"#2563EB", teal:"#0D9488",
};

var TYPE_INFO = {
  volunteer:  {label:"מתנדב",       bg:"#DBEAFE",col:C.blue,  icon:"🙋"},
  manager:    {label:"אחראי משמרת", bg:"#EDE9FE",col:C.purple,icon:"⭐"},
  day_manager:{label:"אחראי יום",   bg:"#CCFBF1",col:C.teal,  icon:"📋"},
  admin:      {label:"מנהל",        bg:"#FEF3C7",col:C.amber, icon:"🛠"},
  superadmin: {label:"מנהל ראשי",   bg:"#FDE8FF",col:"#9333EA",icon:"👑"},
};

// ─── Helpers ────────────────────────────────────────────────────────────────
var _sid = 1000;
function makeShiftId(day) {
  _sid = _sid + 1;
  return day + "-x" + _sid;
}

function flatShifts(shiftMap) {
  var result = [];
  var dayNums = Object.keys(shiftMap);
  for (var di = 0; di < dayNums.length; di++) {
    var d = Number(dayNums[di]);
    var arr = shiftMap[d] || [];
    for (var si = 0; si < arr.length; si++) {
      var s = arr[si];
      result.push({
        id: s.id, day: d, name: s.name, hours: s.hours,
        icon: s.icon, maxVolunteers: s.maxVol, maxManagers: s.maxMgr,
      });
    }
  }
  return result;
}

function getOccupancy(shifts, users, regs) {
  var occ = {};
  for (var i = 0; i < shifts.length; i++) {
    occ[shifts[i].id] = {volunteers:[], managers:[]};
  }
  var codes = Object.keys(regs);
  for (var j = 0; j < codes.length; j++) {
    var code = codes[j];
    var sid = regs[code];
    var u = users[code];
    if (!occ[sid] || !u) continue;
    if (u.type === "volunteer") occ[sid].volunteers.push(code);
    if (u.type === "manager")   occ[sid].managers.push(code);
  }
  return occ;
}

function getDayMgrOcc(dmRegs) {
  var occ = {};
  for (var d = 1; d <= 10; d++) occ[d] = [];
  var codes = Object.keys(dmRegs);
  for (var i = 0; i < codes.length; i++) {
    var day = dmRegs[codes[i]];
    if (occ[day]) occ[day].push(codes[i]);
  }
  return occ;
}

var _logId = 1;
function makeLog(type, uid, uname, shift, day, dayNames, actor) {
  var dayLabel = shift ? (dayNames[shift.day] || ("יום " + shift.day))
                       : day ? (dayNames[day]  || ("יום " + day)) : null;
  return {
    id: _logId++, type: type, userId: uid, userName: uname,
    shiftId: shift ? shift.id : null,
    shiftName: shift ? shift.name : null,
    shiftHours: shift ? shift.hours : null,
    dayLabel: dayLabel,
    actorId: actor.id, actorName: actor.name, actorType: actor.type,
    ts: new Date().toISOString(),
  };
}

function fmtTime(iso) {
  return new Date(iso).toLocaleString("he-IL",{
    day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"
  });
}

function colp(a, b) { return a + b; }

// ─── Export to Excel ─────────────────────────────────────────────────────────
function doExport(users, shifts, regs, dmRegs, dayNames) {
  var rows = [];
  var rkeys = Object.keys(regs);
  for (var i = 0; i < rkeys.length; i++) {
    var code = rkeys[i];
    var u = users[code] || {};
    var s = null;
    for (var j = 0; j < shifts.length; j++) {
      if (shifts[j].id === regs[code]) { s = shifts[j]; break; }
    }
    rows.push({
      "שם": u.name || "", "ת.ז.": code,
      "סוג": u.type === "manager" ? "אחראי משמרת" : "מתנדב",
      "טלפון": u.phone || "", "דוא\"ל": u.email || "",
      "יום": s ? (dayNames[s.day] || ("יום " + s.day)) : "",
      "משמרת": s ? s.name : "", "שעות": s ? s.hours : "",
    });
  }
  var dmkeys = Object.keys(dmRegs);
  for (var k = 0; k < dmkeys.length; k++) {
    var dcode = dmkeys[k];
    var du = users[dcode] || {};
    var dayNum = dmRegs[dcode];
    rows.push({
      "שם": du.name || "", "ת.ז.": dcode, "סוג": "אחראי יום",
      "טלפון": du.phone || "", "דוא\"ל": du.email || "",
      "יום": dayNames[dayNum] || ("יום " + dayNum), "משמרת": "-", "שעות": "-",
    });
  }
  var ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [22,14,16,16,26,18,10,14].map(function(w){return {wch:w};});
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "רישומים");
  XLSX.writeFile(wb, "רישום_משמרות.xlsx");
}

function doExportLog(log, dayNames) {
  var rows = log.map(function(e) {
    return {
      "זמן": fmtTime(e.ts),
      "פעולה": e.type === "register" ? "רישום" : e.type === "remove" ? "הסרה" : "מחיקה",
      "משתמש": e.userName, "ת.ז.": e.userId,
      "יום": e.dayLabel || "-", "משמרת": e.shiftName || "-",
      "בוצע ע\"י": e.actorName, "תפקיד": e.actorType,
    };
  });
  var ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [18,12,20,14,18,10,20,14].map(function(w){return {wch:w};});
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "לוג");
  XLSX.writeFile(wb, "לוג_פעילות.xlsx");
}

function parseImport(file, onDone, onError) {
  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var wb   = XLSX.read(e.target.result, {type:"binary"});
      var ws   = wb.Sheets[wb.SheetNames[0]];
      var rows = XLSX.utils.sheet_to_json(ws, {defval:""});
      var TYPE_MAP = {
        "volunteer":"volunteer","מתנדב":"volunteer","מתנדבת":"volunteer",
        "manager":"manager","אחראי":"manager","אחראי משמרת":"manager",
        "day_manager":"day_manager","אחראי יום":"day_manager",
        "admin":"admin","מנהל":"admin",
        "superadmin":"superadmin","מנהל ראשי":"superadmin",
      };
      function col(row) {
        var keys = ["id","ת.ז.","ת.ז","תעודת זהות","code","קוד","ID"];
        for (var i = 0; i < keys.length; i++) {
          var v = row[keys[i]];
          if (v !== undefined && String(v).trim() !== "") return String(v).trim();
        }
        return "";
      }
      function colK(row, k) {
        var v = row[k];
        return (v !== undefined && String(v).trim() !== "") ? String(v).trim() : "";
      }
      var users = {}, errors = [];
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        var id   = col(row);
        var type = TYPE_MAP[(colK(row,"type") || colK(row,"סוג")).toLowerCase()] || "";
        var name = colK(row,"name") || colK(row,"שם");
        var phone= colK(row,"phone") || colK(row,"טלפון");
        var email= colK(row,"email") || colK(row,"דוא\"ל") || colK(row,"EMAIL");
        var pass = colK(row,"password") || colK(row,"סיסמה");
        var hr   = colK(row,"hr") || colK(row,"HR") || colK(row,"הערות HR") || colK(row,"הערות");
        if (!id)   { errors.push("שורה " + (i+2) + ": ת.ז. חסרה"); continue; }
        if (!type) { errors.push("שורה " + (i+2) + " (" + id + "): סוג לא מוכר"); continue; }
        if (!name) { errors.push("שורה " + (i+2) + " (" + id + "): שם חסר"); continue; }
        var obj = {type:type, name:name, phone:phone, email:email};
        if (pass) obj.password = pass;
        if (hr)   obj.hr = hr;
        users[id] = obj;
      }
      if (!Object.keys(users).length && errors.length) onError(errors);
      else onDone(users, errors);
    } catch(err) {
      onError(["שגיאה: " + err.message]);
    }
  };
  reader.readAsBinaryString(file);
}

// ═══════════════════════════════════════════════════════════════════════════
//  APP
// ═══════════════════════════════════════════════════════════════════════════
// --- DB: load everything on startup ------------------------------------------
async function loadAll() {
  var r = await Promise.all([
    supabase.from("users").select("*"),
    supabase.from("shift_defs").select("*").order("day").order("sort_order"),
    supabase.from("registrations").select("*"),
    supabase.from("day_manager_regs").select("*"),
    supabase.from("day_configs").select("*").order("day"),
    supabase.from("app_settings").select("*"),
    supabase.from("activity_log").select("*").order("created_at",{ascending:false}).limit(500),
  ]);
  var usersRes=r[0],shiftsRes=r[1],regsRes=r[2],dmRes=r[3],dcRes=r[4],settRes=r[5],logRes=r[6];
  if(usersRes.error) throw usersRes.error;

  var users={};
  (usersRes.data||[]).forEach(function(u){
    users[u.id]={type:u.type,name:u.name,phone:u.phone,email:u.email,password:u.password,hr:u.hr||null};
  });

  var shiftMap={};
  (shiftsRes.data||[]).forEach(function(s){
    if(!shiftMap[s.day]) shiftMap[s.day]=[];
    shiftMap[s.day].push({id:s.id,name:s.name,hours:s.hours,icon:s.icon,maxVol:s.max_volunteers,maxMgr:s.max_managers,sortOrder:s.sort_order});
  });

  var regs={};
  (regsRes.data||[]).forEach(function(r){ regs[r.user_id]=r.shift_id; });

  var dmRegs={};
  (dmRes.data||[]).forEach(function(r){ dmRegs[r.user_id]=r.day; });

  var dayConfigs={},dayNames={};
  (dcRes.data||[]).forEach(function(d){
    dayConfigs[d.day]={maxDayMgr:d.max_day_managers};
    dayNames[d.day]=d.name;
  });

  var regOpen=false,allowSelfRemove=true;
  (settRes.data||[]).forEach(function(s){
    if(s.key==="registration_open") regOpen=s.value==="true";
    if(s.key==="allow_self_remove") allowSelfRemove=s.value!=="false";
  });

  var log=(logRes.data||[]).map(function(e){
    return{id:e.id,type:e.type,userId:e.user_id,userName:e.user_name,
      shiftId:e.shift_id,shiftName:e.shift_name,shiftHours:e.shift_hours,
      dayLabel:e.day_label,actorId:e.actor_id,actorName:e.actor_name,actorType:e.actor_type,ts:e.created_at};
  });

  return{users,shiftMap,regs,dmRegs,dayConfigs,dayNames,regOpen,allowSelfRemove,log};
}

// --- Monday.com sync --------------------------------------------------------
var MONDAY_BOARD_ID = "18419606261";
var MONDAY_API_KEY  = "REPLACE_WITH_NEW_API_KEY"; // החלף ב-API key החדש שלך

function mondayQuery(query) {
  return fetch("https://api.monday.com/v2", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": MONDAY_API_KEY,
    },
    body: JSON.stringify({query: query}),
  }).then(function(r){ return r.json(); });
}

// Find the Monday item ID where the "name" column matches the given ID number
function findMondayItem(idNumber) {
  var q = '{ boards(ids: ' + MONDAY_BOARD_ID + ') { items_page(limit: 500) { items { id name column_values(ids: ["name"]) { text } } } } }';
  return mondayQuery(q).then(function(data) {
    try {
      var items = data.data.boards[0].items_page.items;
      for (var i = 0; i < items.length; i++) {
        var item = items[i];
        // Try column_values first, fall back to item name
        var nameVal = "";
        if (item.column_values && item.column_values[0]) {
          nameVal = item.column_values[0].text || "";
        }
        if (!nameVal) nameVal = item.name || "";
        if (nameVal === idNumber || nameVal.trim() === idNumber.trim()) {
          return item.id;
        }
      }
      console.warn("Monday: no item found for ID", idNumber);
      return null;
    } catch(err) {
      console.error("Monday findMondayItem error:", err, JSON.stringify(data));
      return null;
    }
  });
}

// Update Monday item columns
function syncMonday(userId, status, dayLabel, shiftHours) {
  findMondayItem(userId).then(function(itemId) {
    if (!itemId) { console.warn("Monday: item not found for", userId); return; }

    // status: "0"=לא קיים, "1"=קיים, "2"=נמחק
    var colValues = {};
    colValues["color_mm4qvjcs"] = {label: status === "0" ? "לא קיים" : status === "1" ? "קיים" : "נמחק"};
    if (dayLabel  !== undefined) colValues["text_mm4qxbn0"] = dayLabel  || "";
    if (shiftHours !== undefined) colValues["text_mm4qsdfw"] = shiftHours || "";

    var colValStr = JSON.stringify(JSON.stringify(colValues));
    var mutation = 'mutation { change_multiple_column_values(board_id: ' + MONDAY_BOARD_ID + ', item_id: ' + itemId + ', column_values: ' + colValStr + ') { id } }';
    mondayQuery(mutation).then(function(res) {
      if (res.errors) console.error("Monday sync error:", res.errors);
    });
  });
}

function dbLog(entry){
  supabase.from("activity_log").insert({
    type:entry.type,user_id:entry.userId,user_name:entry.userName,
    shift_id:entry.shiftId,shift_name:entry.shiftName,shift_hours:entry.shiftHours,
    day_label:entry.dayLabel,actor_id:entry.actorId,actor_name:entry.actorName,actor_type:entry.actorType,
  }).then(function(res){
    if(res.error) console.error("LOG FAILED:", res.error.code, res.error.message);
  });
}

function pushLog(setLog,entry){
  setLog(function(p){ return [Object.assign({ts:new Date().toISOString(),id:Date.now()},entry)].concat(p); });
}

// ===========================================================================
//  APP ROOT
// ===========================================================================
export default function App() {
  var s0=useState(true);  var loading=s0[0];  var setLoading=s0[1];
  var s1=useState(null);  var loadErr=s1[0];  var setLoadErr=s1[1];
  var s2=useState({});    var users=s2[0];    var setUsers=s2[1];
  var s3=useState({});    var regs=s3[0];     var setRegs=s3[1];
  var s4=useState({});    var dmRegs=s4[0];   var setDmRegs=s4[1];
  var s5=useState({});    var shiftMap=s5[0]; var setShiftMap=s5[1];
  var s6=useState({});    var dayNames=s6[0]; var setDayNames=s6[1];
  var s7=useState({});    var dayConfs=s7[0]; var setDayConfs=s7[1];
  var s8=useState([]);    var log=s8[0];      var setLog=s8[1];
  var s9=useState(false); var regOpen=s9[0];  var setRegOpen=s9[1];
  var s10=useState(true); var allowSelfRemove=s10[0]; var setAllowSelfRemove=s10[1];
  var s11=useState(null); var me=s11[0];      var setMe=s11[1];
  var s12=useState("");   var loginErr=s12[0];var setLoginErr=s12[1];

  var shifts = useMemo(function(){ return flatShifts(shiftMap); },[shiftMap]);
  var occ    = useMemo(function(){ return getOccupancy(shifts,users,regs); },[shifts,users,regs]);
  var dmOcc  = useMemo(function(){ return getDayMgrOcc(dmRegs); },[dmRegs]);

  useEffect(function(){
    loadAll()
      .then(function(data){
        setUsers(data.users); setShiftMap(data.shiftMap);
        setRegs(data.regs); setDmRegs(data.dmRegs);
        setDayConfs(data.dayConfigs); setDayNames(data.dayNames);
        setRegOpen(data.regOpen); setAllowSelfRemove(data.allowSelfRemove);
        setLog(data.log); setLoading(false);
      })
      .catch(function(err){ setLoadErr("שגיאה: "+err.message); setLoading(false); });
  },[]);

  useEffect(function(){
    var ch=supabase.channel("live")
      .on("postgres_changes",{event:"*",schema:"public",table:"registrations"},function(p){
        if(p.eventType==="INSERT") setRegs(function(prev){var n=Object.assign({},prev);n[p.new.user_id]=p.new.shift_id;return n;});
        if(p.eventType==="DELETE") setRegs(function(prev){var n=Object.assign({},prev);delete n[p.old.user_id];return n;});
      })
      .on("postgres_changes",{event:"*",schema:"public",table:"day_manager_regs"},function(p){
        if(p.eventType==="INSERT") setDmRegs(function(prev){var n=Object.assign({},prev);n[p.new.user_id]=p.new.day;return n;});
        if(p.eventType==="DELETE") setDmRegs(function(prev){var n=Object.assign({},prev);delete n[p.old.user_id];return n;});
      })
      .on("postgres_changes",{event:"UPDATE",schema:"public",table:"app_settings"},function(p){
        if(p.new.key==="registration_open") setRegOpen(p.new.value==="true");
        if(p.new.key==="allow_self_remove") setAllowSelfRemove(p.new.value!=="false");
      })
      .subscribe();
    return function(){ supabase.removeChannel(ch); };
  },[]);

  function handleLogin(id,pass){
    var key=id.trim(),u=users[key];
    if(!u){setLoginErr("מספר ת.ז. לא תקין.");return;}
    if(u.type==="admin"||u.type==="superadmin"){if(pass!==u.password){setLoginErr("סיסמה שגויה.");return;}}
    setMe(Object.assign({id:key},u)); setLoginErr("");
  }

  function handleRegister(shiftId){
    supabase.rpc("register_for_shift",{p_user_id:me.id,p_shift_id:shiftId}).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      if(!res.data.success){var m={shift_full:"המשמרת מלאה.",already_registered:"כבר רשום/ה.",registration_closed:"ההרשמה סגורה."};alert(m[res.data.error]||res.data.error);return;}
      setRegs(function(p){var n=Object.assign({},p);n[me.id]=shiftId;return n;});
      var shift=null;for(var i=0;i<shifts.length;i++){if(shifts[i].id===shiftId){shift=shifts[i];break;}}
      var e={type:"register",userId:me.id,userName:me.name,shiftId:shiftId,shiftName:shift?shift.name:null,shiftHours:shift?shift.hours:null,dayLabel:shift?(dayNames[shift.day]||("יום "+shift.day)):null,actorId:me.id,actorName:me.name,actorType:me.type};
      dbLog(e); pushLog(setLog,e);
      syncMonday(me.id, "1", shift?(dayNames[shift.day]||("יום "+shift.day)):"", shift?shift.hours:"");
    });
  }

  function handleAdminDmRegister(userId, day) {
    var dayConfig = dayConfs[day] || {maxDayMgr:2};
    if ((dmOcc[day]||[]).length >= dayConfig.maxDayMgr) { alert("היום מלא — אין מקום לאחראי יום נוסף."); return; }
    if (dmRegs[userId]) { alert("המשתמש כבר רשום ליום אחר."); return; }
    supabase.from("day_manager_regs").insert({user_id:userId, day:day}).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setDmRegs(function(p){var n=Object.assign({},p);n[userId]=day;return n;});
      var e={type:"register",userId:userId,userName:(users[userId]||{}).name||userId,shiftId:null,shiftName:null,shiftHours:null,dayLabel:dayNames[day]||("יום "+day),actorId:me.id,actorName:me.name,actorType:me.type};
      dbLog(e); pushLog(setLog,e);
      syncMonday(userId, "1", dayNames[day]||("יום "+day), "אחראי יום");
    });
  }

  function handleAdminRegister(userId,shiftId){
    var shift=null;for(var i=0;i<shifts.length;i++){if(shifts[i].id===shiftId){shift=shifts[i];break;}}
    if (regs[userId]) { alert("המשתמש כבר רשום למשמרת אחרת."); return; }
    if (shift) {
      var o = occ[shiftId]||{volunteers:[],managers:[]};
      var u = users[userId];
      if (u && u.type==="volunteer" && o.volunteers.length >= shift.maxVolunteers) { alert("המשמרת מלאה."); return; }
      if (u && u.type==="manager"   && o.managers.length   >= shift.maxManagers)   { alert("אין מקום לאחראי משמרת נוסף."); return; }
    }
    supabase.from("registrations").insert({user_id:userId, shift_id:shiftId}).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setRegs(function(p){var n=Object.assign({},p);n[userId]=shiftId;return n;});
      var e={type:"register",userId:userId,userName:(users[userId]||{}).name||userId,shiftId:shiftId,shiftName:shift?shift.name:null,shiftHours:shift?shift.hours:null,dayLabel:shift?(dayNames[shift.day]||("יום "+shift.day)):null,actorId:me.id,actorName:me.name,actorType:me.type};
      dbLog(e); pushLog(setLog,e);
      syncMonday(userId, "1", shift?(dayNames[shift.day]||("יום "+shift.day)):"", shift?shift.hours:"");
    });
  }

  function handleDmRegister(day){
    supabase.rpc("register_day_manager",{p_user_id:me.id,p_day:day}).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      if(!res.data.success){var m={day_full:"היום מלא.",already_registered:"כבר רשום/ה.",registration_closed:"ההרשמה סגורה."};alert(m[res.data.error]||res.data.error);return;}
      setDmRegs(function(p){var n=Object.assign({},p);n[me.id]=day;return n;});
      var e={type:"register",userId:me.id,userName:me.name,shiftId:null,shiftName:null,shiftHours:null,dayLabel:dayNames[day]||("יום "+day),actorId:me.id,actorName:me.name,actorType:me.type};
      dbLog(e); pushLog(setLog,e);
      syncMonday(me.id, "1", dayNames[day]||("יום "+day), "אחראי יום");
    });
  }

  function handleRemove(uid){
    var sid=regs[uid];
    var shift=null;for(var i=0;i<shifts.length;i++){if(shifts[i].id===sid){shift=shifts[i];break;}}
    supabase.from("registrations").delete().eq("user_id",uid).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setRegs(function(p){var n=Object.assign({},p);delete n[uid];return n;});
      var e={type:"remove",userId:uid,userName:(users[uid]||{}).name||uid,shiftId:sid,shiftName:shift?shift.name:null,shiftHours:shift?shift.hours:null,dayLabel:shift?(dayNames[shift.day]||("יום "+shift.day)):null,actorId:me.id,actorName:me.name,actorType:me.type};
      dbLog(e); pushLog(setLog,e);
      syncMonday(uid, "0", "", "");
    });
  }

  function handleDmRemove(uid){
    var day=dmRegs[uid];
    supabase.from("day_manager_regs").delete().eq("user_id",uid).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setDmRegs(function(p){var n=Object.assign({},p);delete n[uid];return n;});
      var e={type:"remove",userId:uid,userName:(users[uid]||{}).name||uid,shiftId:null,shiftName:null,shiftHours:null,dayLabel:dayNames[day]||("יום "+day),actorId:me.id,actorName:me.name,actorType:me.type};
      dbLog(e); pushLog(setLog,e);
      syncMonday(uid, "0", "", "");
    });
  }

  function handleDeleteUser(uid){
    supabase.from("users").delete().eq("id",uid).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setUsers(function(p){var n=Object.assign({},p);delete n[uid];return n;});
      var e={type:"delete_user",userId:uid,userName:(users[uid]||{}).name||uid,shiftId:null,shiftName:null,shiftHours:null,dayLabel:null,actorId:me.id,actorName:me.name,actorType:me.type};
      dbLog(e); pushLog(setLog,e);
      syncMonday(uid, "2", "", "");
    });
  }

  function handleImport(newArr){
    var cleaned = newArr.map(function(u){
      var obj = {id:u.id, type:u.type, name:u.name};
      if(u.phone)    obj.phone    = u.phone;
      if(u.email)    obj.email    = u.email;
      if(u.password) obj.password = String(u.password);
      if(u.hr)       obj.hr       = u.hr;
      return obj;
    });
    supabase.from("users").upsert(cleaned,{onConflict:"id"}).then(function(res){
      if(res.error){alert("שגיאה בייבוא: "+res.error.message);return;}
      setUsers(function(prev){
        var n=Object.assign({},prev);
        for(var i=0;i<newArr.length;i++){
          var u=newArr[i];
          n[u.id]={type:u.type,name:u.name,phone:u.phone,email:u.email,password:u.password||null,hr:u.hr||null};
        }
        return n;
      });
      // Log each imported user — skip if nothing changed
      var logEntries = [];
      newArr.forEach(function(u){
        var id=u.id; var existing=users[id];
        var logType;
        if (!existing) {
          logType = "import_new";
        } else if (existing.type !== u.type) {
          logType = "import_role";
        } else {
          // Check if any detail actually changed
          var changed = existing.name !== u.name ||
            (existing.phone||"") !== (u.phone||"") ||
            (existing.email||"") !== (u.email||"") ||
            (existing.hr||"") !== (u.hr||"") ||
            (existing.password||"") !== (u.password||"");
          if (!changed) return; // nothing changed — skip log
          logType = "import_update";
        }
        logEntries.push({type:logType,user_id:id,user_name:u.name,actor_id:me.id,actor_name:me.name,actor_type:me.type});
      });
      logEntries.forEach(function(e){ supabase.from("activity_log").insert(e); });
      setLog(function(prev){
        var newEntries = logEntries.map(function(e){
          return{id:Date.now()+Math.random(),type:e.type,userId:e.user_id,userName:e.user_name,
            shiftId:null,shiftName:null,shiftHours:null,dayLabel:null,
            actorId:e.actor_id,actorName:e.actor_name,actorType:e.actor_type,ts:new Date().toISOString()};
        });
        return newEntries.concat(prev);
      });
    });
  }

  function handleUpdateShift(day,shiftId,field,value){
    var dbField=field==="maxVol"?"max_volunteers":field==="maxMgr"?"max_managers":field;
    var upd={}; upd[dbField]=value;
    supabase.from("shift_defs").update(upd).eq("id",shiftId).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setShiftMap(function(prev){var next=Object.assign({},prev);next[day]=(prev[day]||[]).map(function(s){if(s.id!==shiftId)return s;var ns=Object.assign({},s);ns[field]=value;return ns;});return next;});
    });
  }

  function handleAddShift(day){
    var newId=makeShiftId(day);
    var maxOrder=0;var arr=shiftMap[day]||[];
    for(var i=0;i<arr.length;i++) if(arr[i].sortOrder>maxOrder) maxOrder=arr[i].sortOrder;
    supabase.from("shift_defs").insert({id:newId,day:day,name:"משמרת חדשה",hours:"00:00-00:00",icon:"⭐",max_volunteers:10,max_managers:1,sort_order:maxOrder+1}).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setShiftMap(function(prev){var next=Object.assign({},prev);next[day]=(prev[day]||[]).concat([{id:newId,name:"משמרת חדשה",hours:"00:00-00:00",icon:"⭐",maxVol:10,maxMgr:1,sortOrder:maxOrder+1}]);return next;});
    });
  }

  function handleRemoveShift(day,shiftId){
    supabase.from("shift_defs").delete().eq("id",shiftId).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setShiftMap(function(prev){var next=Object.assign({},prev);next[day]=(prev[day]||[]).filter(function(s){return s.id!==shiftId;});return next;});
      setRegs(function(prev){var n=Object.assign({},prev);var keys=Object.keys(n);for(var i=0;i<keys.length;i++){if(n[keys[i]]===shiftId)delete n[keys[i]];}return n;});
    });
  }

  function handleUpdateDayName(day,name){
    supabase.from("day_configs").update({name:name}).eq("day",day).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setDayNames(function(p){var n=Object.assign({},p);n[day]=name;return n;});
    });
  }

  function handleUpdateDayConfig(day,val){
    var v=Math.min(2,Math.max(0,Number(val)));
    supabase.from("day_configs").update({max_day_managers:v}).eq("day",day).then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setDayConfs(function(p){var n=Object.assign({},p);n[day]={maxDayMgr:v};return n;});
    });
  }

  function handleToggleReg(){
    var v=regOpen?"false":"true";
    supabase.from("app_settings").update({value:v}).eq("key","registration_open").then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setRegOpen(!regOpen);
    });
  }

  function handleToggleSelfRemove(){
    var v=allowSelfRemove?"false":"true";
    supabase.from("app_settings").update({value:v}).eq("key","allow_self_remove").then(function(res){
      if(res.error){alert("שגיאה: "+res.error.message);return;}
      setAllowSelfRemove(!allowSelfRemove);
    });
  }

  if(loading){
    return (
      <div dir="rtl" style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Segoe UI',Arial,sans-serif"}}>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:40,marginBottom:16}}>&#9203;</div>
          <div style={{fontSize:18,fontWeight:700,color:C.navy}}>טוען נתונים...</div>
        </div>
      </div>
    );
  }

  if(loadErr){
    return (
      <div dir="rtl" style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
        <div style={{background:C.card,borderRadius:16,padding:32,maxWidth:400,textAlign:"center",boxShadow:"0 4px 20px rgba(0,0,0,.1)"}}>
          <div style={{fontSize:36,marginBottom:12}}>&#10060;</div>
          <div style={{fontSize:16,fontWeight:700,color:C.red,marginBottom:8}}>שגיאת חיבור</div>
          <div style={{fontSize:13,color:C.muted,marginBottom:16}}>{loadErr}</div>
          <div style={{fontSize:12,color:C.muted}}>בדוק ש-VITE_SUPABASE_URL ו-VITE_SUPABASE_ANON_KEY מוגדרים נכון.</div>
        </div>
      </div>
    );
  }

  if(!me) return <LoginScreen onLogin={handleLogin} error={loginErr} users={users} />;

  var isSup=me.type==="superadmin";
  var isAdm=me.type==="admin"||isSup;
  var isDm=me.type==="day_manager";

  if(isAdm) return (
    <AdminPanel
      me={me} users={users} shifts={shifts} shiftMap={shiftMap}
      regs={regs} dmRegs={dmRegs} occ={occ} dmOcc={dmOcc}
      dayNames={dayNames} dayConfigs={dayConfs} log={log}
      isSup={isSup} regOpen={regOpen} allowSelfRemove={allowSelfRemove}
      onToggleReg={handleToggleReg} onToggleSelfRemove={handleToggleSelfRemove}
      onRemove={handleRemove} onDmRemove={handleDmRemove}
      onDeleteUser={handleDeleteUser} onImport={handleImport}
      onAdminRegister={handleAdminRegister}
      onAdminDmRegister={handleAdminDmRegister}
      onUpdateShift={handleUpdateShift} onAddShift={handleAddShift} onRemoveShift={handleRemoveShift}
      onUpdateDayName={handleUpdateDayName} onUpdateDayConfig={handleUpdateDayConfig}
      onLogout={function(){setMe(null);setLoginErr("");}}
    />
  );

  if(isDm) return (
    <DayMgrView
      me={me} dayNames={dayNames} dmRegs={dmRegs} dmOcc={dmOcc}
      dayConfigs={dayConfs} regOpen={regOpen} allowSelfRemove={allowSelfRemove}
      shifts={shifts} occ={occ} users={users}
      onRegister={handleDmRegister}
      onSelfRemove={function(){handleDmRemove(me.id);}}
      onLogout={function(){setMe(null);setLoginErr("");}}
    />
  );

  return (
    <VolView
      me={me} shifts={shifts} regs={regs} occ={occ}
      dayNames={dayNames} regOpen={regOpen} allowSelfRemove={allowSelfRemove}
      users={users} dmOcc={dmOcc} dmRegs={dmRegs}
      onRegister={handleRegister}
      onSelfRemove={function(){handleRemove(me.id);}}
      onLogout={function(){setMe(null);setLoginErr("");}}
    />
  );
}

function Hdr(props) {
  return (
    <header dir="rtl" style={{background:props.bg||C.navy,color:"#fff",padding:"0 20px",height:60,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:99,boxShadow:"0 2px 10px rgba(0,0,0,.3)"}}>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <img src="/logo.png" alt="logo" style={{width:36,height:36,objectFit:"contain",borderRadius:4}} />
        <div>
          <div style={{fontSize:11,color:"#C4B5FD",fontWeight:700,lineHeight:1}}>מידברן 2026 - מחלקת תנועה</div>
          <div style={{fontWeight:800,fontSize:15,lineHeight:1.3}}>{props.title}</div>
        </div>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <div style={{textAlign:"left"}}>
          <div style={{fontSize:13,fontWeight:700}}>{props.name}</div>
          <div style={{fontSize:11,color:"#93C5FD"}}>{props.sub}</div>
        </div>
        <button onClick={props.onLogout} style={{background:"rgba(255,255,255,.12)",border:"1px solid rgba(255,255,255,.25)",color:"#fff",borderRadius:8,padding:"5px 13px",cursor:"pointer",fontSize:12,fontWeight:600}}>יציאה</button>
      </div>
    </header>
  );
}

function DayPill(props) {
  var active = props.active;
  return (
    <button onClick={props.onClick} style={{padding:"5px 14px",borderRadius:20,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,background:active?C.navy:"#fff",color:active?"#fff":C.navy,boxShadow:active?"0 3px 10px rgba(15,45,74,.35)":"0 1px 4px rgba(0,0,0,.1)"}}>
      {props.label}
    </button>
  );
}

function Bar(props) {
  var pct = props.max > 0 ? Math.min(1, props.val / props.max) : 0;
  var col = pct >= 1 ? C.red : pct >= 0.8 ? C.amber : C.green;
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}>
        <span style={{color:C.muted}}>{props.label || "מקומות"}</span>
        <span style={{fontWeight:700,color:pct>=1?C.red:C.text}}>{props.max - props.val} / {props.max} פנויים</span>
      </div>
      <div style={{height:6,background:"#E2E8F0",borderRadius:3,overflow:"hidden"}}>
        <div style={{height:"100%",borderRadius:3,width:(pct*100)+"%",background:col}} />
      </div>
    </div>
  );
}

// ─── LOGIN ───────────────────────────────────────────────────────────────────
function LoginScreen(props) {
  var si = useState(""); var idVal = si[0]; var setIdVal = si[1];
  var sp = useState(""); var passVal = sp[0]; var setPassVal = sp[1];
  var ss = useState("id"); var step = ss[0]; var setStep = ss[1];
  var sdt = useState(null); var detType = sdt[0]; var setDetType = sdt[1];

  function next() {
    var key = idVal.trim();
    var u = props.users[key];
    if (!u) { props.onLogin(key, ""); return; }
    if (u.type === "admin" || u.type === "superadmin") {
      setDetType(u.type); setStep("password");
    } else {
      props.onLogin(key, "");
    }
  }

  var samples = [
    {id:"123456789",label:"מתנדב"},
    {id:"891234567",label:"אחראי יום"},
    {id:"567890123",label:"אחראי משמרת"},
    {id:"111111111",label:"מנהל (Admin123!)"},
    {id:"000000001",label:"מנהל ראשי (Shift2025!)"},
  ];

  return (
    <div dir="rtl" style={{minHeight:"100vh",background:"url('/login-bg.jpg') center center / cover no-repeat",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Segoe UI',Arial,sans-serif",padding:20}}>
      <div style={{background:"rgba(255,255,255,0.92)",backdropFilter:"blur(8px)",borderRadius:20,padding:"36px 32px",width:"100%",maxWidth:390,boxShadow:"0 30px 80px rgba(0,0,0,.5)"}}>
        <div style={{textAlign:"center",marginBottom:26}}>
          <img src="/logo.png" alt="logo" style={{width:90,height:90,objectFit:"contain",margin:"0 auto 12px",display:"block"}} />
          <div style={{color:"#7C3AED",fontSize:16,fontWeight:800,marginBottom:4}}>מידברן 2026 - מחלקת תנועה</div>
          <h2 style={{color:C.navy,margin:"0 0 4px",fontSize:24,fontWeight:800}}>רישום משמרות</h2>
          <p style={{color:C.muted,margin:0,fontSize:13}}>{step==="id"?"הזן/י מספר תעודת זהות":"הזן/י סיסמה להמשך"}</p>
        </div>

        {step === "id" && (
          <div>
            <input value={idVal} onChange={function(e){setIdVal(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")next();}} placeholder="000000000" autoFocus
              style={{width:"100%",padding:"13px 16px",borderRadius:10,border:"2.5px solid " + (props.error?C.red:"#CBD5E0"),fontSize:22,textAlign:"center",letterSpacing:4,outline:"none",boxSizing:"border-box",direction:"ltr",fontFamily:"monospace",color:"#1A202C",background:"#fff",fontWeight:700}} />
            {props.error && <div style={{background:"#FEF2F2",border:"1px solid "+C.red,borderRadius:8,padding:"7px 12px",marginTop:8,color:C.red,fontSize:13,textAlign:"center"}}>{props.error}</div>}
            <button onClick={next} style={{width:"100%",marginTop:14,padding:"13px 0",background:"linear-gradient(135deg,#E67E22,#F39C12)",color:"#fff",border:"none",borderRadius:10,fontSize:16,fontWeight:800,cursor:"pointer"}}>המשך</button>
          </div>
        )}

        {step === "password" && (
          <div>
            <div style={{background:"#F0F4F8",borderRadius:9,padding:"9px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:9}}>
              <span style={{fontSize:18}}>{detType==="superadmin"?"👑":"🛠"}</span>
              <div>
                <div style={{fontSize:13,fontWeight:700,color:C.navy}}>{(props.users[idVal.trim()]||{}).name}</div>
                <div style={{fontSize:11,color:C.muted}}>{(TYPE_INFO[detType]||{}).label} - ת.ז. {idVal.trim()}</div>
              </div>
              <button onClick={function(){setStep("id");setPassVal("");setDetType(null);}} style={{marginRight:"auto",background:"none",border:"none",color:C.muted,fontSize:11,cursor:"pointer",textDecoration:"underline"}}>חזור</button>
            </div>
            <input type="password" value={passVal} onChange={function(e){setPassVal(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")props.onLogin(idVal.trim(),passVal);}} placeholder="סיסמה" autoFocus
              style={{width:"100%",padding:"13px 16px",borderRadius:10,border:"2.5px solid "+(props.error?C.red:"#CBD5E0"),fontSize:18,textAlign:"center",outline:"none",boxSizing:"border-box",color:"#1A202C",background:"#fff",fontWeight:700,letterSpacing:3}} />
            {props.error && <div style={{background:"#FEF2F2",border:"1px solid "+C.red,borderRadius:8,padding:"7px 12px",marginTop:8,color:C.red,fontSize:13,textAlign:"center"}}>{props.error}</div>}
            <button onClick={function(){props.onLogin(idVal.trim(),passVal);}} style={{width:"100%",marginTop:14,padding:"13px 0",background:"linear-gradient(135deg,#E67E22,#F39C12)",color:"#fff",border:"none",borderRadius:10,fontSize:16,fontWeight:800,cursor:"pointer"}}>כניסה</button>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── VOLUNTEER VIEW ───────────────────────────────────────────────────────────
function VolView(props) {
  var sf = useState("all"); var filterDay = sf[0]; var setFilterDay = sf[1];
  var sc = useState(null); var confirming = sc[0]; var setConfirming = sc[1];
  var sv = useState("register"); var view = sv[0]; var setView = sv[1];
  var isMgr = props.me.type === "manager";
  var myShiftId = props.regs[props.me.id] || null;
  var myShift = null;
  for (var i = 0; i < props.shifts.length; i++) {
    if (props.shifts[i].id === myShiftId) { myShift = props.shifts[i]; break; }
  }

  function getSlot(shift) {
    var o = props.occ[shift.id] || {volunteers:[],managers:[]};
    if (isMgr) return {used:o.managers.length, max:shift.maxManagers, ok:o.managers.length<shift.maxManagers};
    return {used:o.volunteers.length, max:shift.maxVolunteers, ok:o.volunteers.length<shift.maxVolunteers};
  }

  var visibleDays = filterDay === "all" ? DAYS : [Number(filterDay)];

  return (
    <div dir="rtl" style={{minHeight:"100vh",background:C.bg,fontFamily:"'Segoe UI',Arial,sans-serif",color:C.text}}>
      <Hdr icon="🗓" title="רישום משמרות" name={props.me.name} sub={(isMgr?"אחראי משמרת":"מתנדב")+" - ת.ז. "+props.me.id} onLogout={props.onLogout} />
      <div style={{maxWidth:900,margin:"0 auto",padding:"24px 16px"}}>

        {/* View toggle */}
        <div style={{display:"flex",gap:4,background:"#fff",borderRadius:10,padding:3,boxShadow:"0 1px 4px rgba(0,0,0,.08)",width:"fit-content",marginBottom:22}}>
          <button onClick={function(){setView("register");}} style={{padding:"7px 18px",borderRadius:7,border:"none",cursor:"pointer",fontSize:13,fontWeight:700,background:view==="register"?C.navy:"transparent",color:view==="register"?"#fff":C.muted}}>
            הרשמה
          </button>
          <button onClick={function(){setView("roster");}} style={{padding:"7px 18px",borderRadius:7,border:"none",cursor:"pointer",fontSize:13,fontWeight:700,background:view==="roster"?C.navy:"transparent",color:view==="roster"?"#fff":C.muted}}>
            📋 רשימת משמרות
          </button>
        </div>

        {view === "roster" && (
          <RosterView shifts={props.shifts} occ={props.occ} users={props.users||{}} dayNames={props.dayNames} dmOcc={props.dmOcc||{}} dmRegs={props.dmRegs||{}} />
        )}

        {view === "register" && myShift && (
          <div style={{background:"linear-gradient(135deg,#27AE60,#1E8449)",color:"#fff",borderRadius:14,padding:"20px 24px",marginBottom:24,boxShadow:"0 4px 20px rgba(39,174,96,.3)"}}>
            <div style={{fontSize:12,opacity:.85,marginBottom:4,fontWeight:600}}>נרשם/ה למשמרת</div>
            <div style={{fontSize:22,fontWeight:800}}>{props.dayNames[myShift.day]||("יום "+myShift.day)} - {myShift.icon} {myShift.name}</div>
            <div style={{fontSize:14,opacity:.85,marginTop:3}}>{myShift.hours}</div>
            {props.allowSelfRemove ? (
              <div style={{marginTop:14,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                <button onClick={props.onSelfRemove} style={{padding:"9px 20px",borderRadius:9,border:"2px solid rgba(255,255,255,.7)",background:"rgba(255,255,255,.15)",color:"#fff",fontSize:13,fontWeight:800,cursor:"pointer"}}>
                  ביטול רישום וחזרה לבחירה
                </button>
                <span style={{fontSize:11,opacity:.75}}>ניתן להירשם מחדש לאחר הביטול</span>
              </div>
            ) : (
              <div style={{marginTop:12,background:"rgba(0,0,0,.18)",borderRadius:8,padding:"8px 12px",fontSize:12}}>רוצה להחליף? יש לפנות למנהל - הוא יסיר אותך ותוכל/י להירשם מחדש.</div>
            )}
          </div>
        )}

        {view === "register" && !props.regOpen && !myShift && (
          <div style={{background:"linear-gradient(135deg,#1A1A2E,#2D2D44)",color:"#fff",borderRadius:14,padding:"26px 28px",marginBottom:24,textAlign:"center"}}>
            <div style={{fontSize:40,marginBottom:10}}>🔒</div>
            <div style={{fontSize:20,fontWeight:900,marginBottom:6}}>ההרשמה למשמרות טרם נפתחה</div>
            <div style={{fontSize:14,opacity:.75}}>ההרשמה תיפתח בקרוב על ידי הנהלת האירוע.</div>
          </div>
        )}

        {view === "register" && !myShift && props.regOpen && (
          <div style={{marginBottom:18}}>
            <h2 style={{color:C.navy,fontSize:20,fontWeight:800,margin:"0 0 4px"}}>בחר/י משמרת</h2>
            <p style={{color:C.muted,fontSize:13,margin:0}}>{isMgr?"בחר/י משמרת להיות אחראי/ת עליה":"בחר/י משמרת להתנדב בה"}</p>
          </div>
        )}

        {view === "register" && (props.regOpen || myShift) && (
          <div>
            <div style={{display:"flex",gap:7,flexWrap:"wrap",marginBottom:20}}>
              <DayPill label="כל הימים" active={filterDay==="all"} onClick={function(){setFilterDay("all");}} />
              {DAYS.map(function(d) {
                return <DayPill key={d} label={props.dayNames[d]||("יום "+d)} active={filterDay===String(d)} onClick={function(){setFilterDay(String(d));}} />;
              })}
            </div>

            {visibleDays.map(function(day) {
              var dayShifts = props.shifts.filter(function(s){ return s.day === day; });
              if (!dayShifts.length) return null;
              return (
                <div key={day} style={{marginBottom:26}}>
                  <div style={{marginBottom:10}}><span style={{background:C.navy,color:"#fff",borderRadius:6,padding:"2px 11px",fontSize:12,fontWeight:700}}>{props.dayNames[day]||("יום "+day)}</span></div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(205px,1fr))",gap:12}}>
                    {dayShifts.map(function(shift) {
                      var slot = getSlot(shift);
                      var isMe = shift.id === myShiftId;
                      var isPending = confirming === shift.id;
                      return (
                        <div key={shift.id} style={{background:C.card,borderRadius:13,padding:16,border:"2.5px solid "+(isMe?C.green:"transparent"),boxShadow:"0 2px 8px rgba(0,0,0,.08)",opacity:(!slot.ok&&!isMe)?0.55:1}}>
                          <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
                            <div>
                              <div style={{fontSize:20}}>{shift.icon}</div>
                              <div style={{fontSize:15,fontWeight:800,color:C.navy,marginTop:2}}>{shift.name}</div>
                              <div style={{fontSize:11,color:C.muted}}>{shift.hours}</div>
                            </div>
                            {isMe && <span style={{background:C.green,color:"#fff",borderRadius:6,padding:"2px 8px",fontSize:10,fontWeight:800,height:"fit-content"}}>שלי</span>}
                          </div>
                          <div style={{marginBottom:12}}>
                            <Bar val={slot.used} max={slot.max} />
                          </div>
                          {!myShiftId && (
                            isPending ? (
                              <div>
                                <div style={{background:"#FEF2F2",border:"2px solid "+C.red,borderRadius:8,padding:"9px 11px",marginBottom:8,textAlign:"center"}}>
                                  <div style={{fontSize:13,fontWeight:900,color:C.red,marginBottom:2}}>שים/י לב!</div>
                                  <div style={{fontSize:11,fontWeight:800,color:C.red}}>לאחר הרישום לא ניתן לשנות את השיבוץ.</div>
                                  <div style={{fontSize:10,color:"#7F1D1D",marginTop:3,fontWeight:700}}>שינויים רק דרך אחראי HR</div>
                                </div>
                                <div style={{fontSize:11,color:C.muted,marginBottom:6,textAlign:"center"}}>לאשר הרשמה?</div>
                                <div style={{display:"flex",gap:6}}>
                                  <button onClick={function(){props.onRegister(shift.id);setConfirming(null);}} style={{flex:1,padding:"7px 0",borderRadius:7,border:"none",background:C.green,color:"#fff",fontSize:12,fontWeight:800,cursor:"pointer"}}>אישור</button>
                                  <button onClick={function(){setConfirming(null);}} style={{flex:1,padding:"7px 0",borderRadius:7,border:"1px solid "+C.muted,background:"transparent",color:C.muted,fontSize:12,cursor:"pointer"}}>ביטול</button>
                                </div>
                              </div>
                            ) : (
                              <button onClick={function(){if(slot.ok)setConfirming(shift.id);}} disabled={!slot.ok}
                                style={{width:"100%",padding:"8px 0",borderRadius:8,border:"none",background:slot.ok?"linear-gradient(135deg,#E67E22,#F39C12)":"#E2E8F0",color:slot.ok?"#fff":C.muted,fontSize:12,fontWeight:800,cursor:slot.ok?"pointer":"not-allowed"}}>
                                {slot.ok?"הירשם/י למשמרת":"המשמרת מלאה"}
                              </button>
                            )
                          )}
                          {myShiftId && !isMe && (
                            <div style={{fontSize:11,color:C.muted,textAlign:"center"}}>{slot.ok?(slot.max-slot.used)+" פנויים":"מלאה"}</div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── DAY MANAGER VIEW ─────────────────────────────────────────────────────────
function DayMgrView(props) {
  var sc = useState(null); var confirming = sc[0]; var setConfirming = sc[1];
  var sv = useState("register"); var view = sv[0]; var setView = sv[1];
  var myDay = props.dmRegs[props.me.id] || null;

  return (
    <div dir="rtl" style={{minHeight:"100vh",background:C.bg,fontFamily:"'Segoe UI',Arial,sans-serif",color:C.text}}>
      <Hdr icon="📋" title="רישום משמרות" name={props.me.name} sub={"אחראי יום - ת.ז. "+props.me.id} onLogout={props.onLogout} />
      <div style={{maxWidth:700,margin:"0 auto",padding:"24px 16px"}}>
        <div style={{display:"flex",gap:4,background:"#fff",borderRadius:10,padding:3,boxShadow:"0 1px 4px rgba(0,0,0,.08)",width:"fit-content",marginBottom:22}}>
          <button onClick={function(){setView("register");}} style={{padding:"7px 18px",borderRadius:7,border:"none",cursor:"pointer",fontSize:13,fontWeight:700,background:view==="register"?C.navy:"transparent",color:view==="register"?"#fff":C.muted}}>הרשמה</button>
          <button onClick={function(){setView("roster");}} style={{padding:"7px 18px",borderRadius:7,border:"none",cursor:"pointer",fontSize:13,fontWeight:700,background:view==="roster"?C.navy:"transparent",color:view==="roster"?"#fff":C.muted}}>📋 רשימת משמרות</button>
        </div>
        {view === "roster" && (
          <RosterView shifts={props.shifts||[]} occ={props.occ||{}} users={props.users||{}} dayNames={props.dayNames} dmOcc={props.dmOcc||{}} dmRegs={props.dmRegs||{}} />
        )}
        {view === "register" && myDay && (
          <div style={{background:"linear-gradient(135deg,#0D9488,#0F766E)",color:"#fff",borderRadius:14,padding:"20px 24px",marginBottom:24}}>
            <div style={{fontSize:12,opacity:.85,marginBottom:4,fontWeight:600}}>רשום/ה כאחראי/ת יום</div>
            <div style={{fontSize:22,fontWeight:800}}>📋 {props.dayNames[myDay]||("יום "+myDay)}</div>
            {props.allowSelfRemove ? (
              <div style={{marginTop:14,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                <button onClick={props.onSelfRemove} style={{padding:"9px 20px",borderRadius:9,border:"2px solid rgba(255,255,255,.7)",background:"rgba(255,255,255,.15)",color:"#fff",fontSize:13,fontWeight:800,cursor:"pointer"}}>
                  ביטול רישום וחזרה לבחירה
                </button>
                <span style={{fontSize:11,opacity:.75}}>ניתן לבחור יום אחר לאחר הביטול</span>
              </div>
            ) : (
              <div style={{marginTop:10,background:"rgba(0,0,0,.18)",borderRadius:8,padding:"7px 11px",fontSize:12}}>רוצה להחליף יום? יש לפנות למנהל.</div>
            )}
          </div>
        )}
        {view === "register" && !props.regOpen && !myDay && (
          <div style={{background:"linear-gradient(135deg,#1A1A2E,#2D2D44)",color:"#fff",borderRadius:14,padding:"26px 28px",marginBottom:24,textAlign:"center"}}>
            <div style={{fontSize:40,marginBottom:10}}>🔒</div>
            <div style={{fontSize:20,fontWeight:900}}>ההרשמה טרם נפתחה</div>
          </div>
        )}
        {view === "register" && !myDay && props.regOpen && <h2 style={{color:C.navy,fontSize:20,fontWeight:800,margin:"0 0 16px"}}>בחר/י יום</h2>}
        {view === "register" && (props.regOpen || myDay) && (
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {DAYS.map(function(day) {
              var occupied = (props.dmOcc[day]||[]).length;
              var maxSlots = (props.dayConfigs[day]||{maxDayMgr:2}).maxDayMgr;
              var ok = occupied < maxSlots;
              var isMe = myDay === day;
              var pct = maxSlots > 0 ? Math.min(1, occupied / maxSlots) : 0;
              var isPending = confirming === day;
              return (
                <div key={day} style={{background:C.card,borderRadius:12,padding:"16px 20px",boxShadow:"0 2px 8px rgba(0,0,0,.08)",border:"2.5px solid "+(isMe?C.teal:"transparent"),opacity:(!ok&&!isMe)?0.55:1}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:14,flexWrap:"wrap"}}>
                    <div style={{display:"flex",alignItems:"center",gap:10}}>
                      <span style={{fontSize:20}}>📋</span>
                      <div>
                        <div style={{fontSize:15,fontWeight:800,color:C.navy}}>{props.dayNames[day]||("יום "+day)}</div>
                        <div style={{fontSize:11,color:C.muted,marginTop:2}}>{occupied}/{maxSlots} אחראי יום</div>
                      </div>
                      {isMe && <span style={{background:C.teal,color:"#fff",borderRadius:6,padding:"2px 8px",fontSize:10,fontWeight:800}}>שלי</span>}
                    </div>
                    <div style={{flex:1,minWidth:100}}>
                      <div style={{height:6,background:"#E2E8F0",borderRadius:3,overflow:"hidden"}}>
                        <div style={{height:"100%",borderRadius:3,width:(pct*100)+"%",background:pct>=1?C.red:pct>=0.5?C.amber:C.teal}} />
                      </div>
                    </div>
                    {!myDay && (
                      isPending ? (
                        <div style={{display:"flex",flexDirection:"column",gap:5,minWidth:160}}>
                          <div style={{background:"#FEF2F2",border:"2px solid "+C.red,borderRadius:7,padding:"7px 10px",textAlign:"center"}}>
                            <div style={{fontSize:12,fontWeight:900,color:C.red}}>לא ניתן לשנות לאחר הרישום!</div>
                            <div style={{fontSize:10,color:"#7F1D1D",fontWeight:700}}>שינויים רק דרך אחראי HR</div>
                          </div>
                          <div style={{display:"flex",gap:5}}>
                            <button onClick={function(){props.onRegister(day);setConfirming(null);}} style={{flex:1,padding:"7px 0",borderRadius:7,border:"none",background:C.teal,color:"#fff",fontSize:12,fontWeight:800,cursor:"pointer"}}>אישור</button>
                            <button onClick={function(){setConfirming(null);}} style={{flex:1,padding:"7px 0",borderRadius:7,border:"1px solid "+C.muted,background:"transparent",color:C.muted,fontSize:12,cursor:"pointer"}}>ביטול</button>
                          </div>
                        </div>
                      ) : (
                        <button onClick={function(){if(ok)setConfirming(day);}} disabled={!ok}
                          style={{padding:"8px 18px",borderRadius:8,border:"none",background:ok?"linear-gradient(135deg,#0D9488,#0F766E)":"#E2E8F0",color:ok?"#fff":C.muted,fontSize:12,fontWeight:800,cursor:ok?"pointer":"not-allowed",whiteSpace:"nowrap"}}>
                          {ok?"הירשם/י":"מקסימום"}
                        </button>
                      )
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── ADMIN PANEL ─────────────────────────────────────────────────────────────
// ─── ROSTER VIEW (names only, visible to all users) ───────────────────────────
function RosterView(props) {
  var sf = useState("all"); var filterDay = sf[0]; var setFilterDay = sf[1];
  var visibleDays = filterDay === "all" ? DAYS : [Number(filterDay)];

  return (
    <div>
      <div style={{marginBottom:16}}>
        <div style={{fontSize:14,color:C.muted,marginBottom:12}}>רשימת הנרשמים לכל משמרת. מוצגים שמות בלבד.</div>
        <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
          <DayPill label="כל הימים" active={filterDay==="all"} onClick={function(){setFilterDay("all");}} />
          {DAYS.map(function(d){
            return <DayPill key={d} label={props.dayNames[d]||("יום "+d)} active={filterDay===String(d)} onClick={function(){setFilterDay(String(d));}} />;
          })}
        </div>
      </div>

      {visibleDays.map(function(day) {
        var dayShifts = props.shifts.filter(function(s){ return s.day === day; });
        if (!dayShifts.length) return null;
        var dayMgrs = (props.dmOcc[day]||[]).map(function(id){ return props.users[id]; }).filter(Boolean);

        return (
          <div key={day} style={{background:C.card,borderRadius:14,marginBottom:18,boxShadow:"0 2px 8px rgba(0,0,0,.08)",overflow:"hidden"}}>
            <div style={{background:C.navy,color:"#fff",padding:"11px 18px",display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:16,fontWeight:900}}>{props.dayNames[day]||("יום "+day)}</span>
              {dayMgrs.length > 0 && (
                <span style={{fontSize:11,opacity:.8}}>
                  אחראי יום: {dayMgrs.map(function(m){return m.name;}).join(", ")}
                </span>
              )}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))"}}>
              {dayShifts.map(function(shift, idx) {
                var o = props.occ[shift.id] || {volunteers:[], managers:[]};
                var volNames = o.volunteers.map(function(id){ return (props.users[id]||{}).name; }).filter(Boolean);
                var mgrNames = o.managers.map(function(id){ return (props.users[id]||{}).name; }).filter(Boolean);
                return (
                  <div key={shift.id} style={{padding:"14px 16px",borderLeft:idx>0?"1px solid #EEF2F7":"none"}}>
                    <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:10,paddingBottom:8,borderBottom:"1.5px solid #EEF2F7"}}>
                      <span style={{fontSize:16}}>{shift.icon}</span>
                      <div>
                        <div style={{fontSize:13,fontWeight:800,color:C.navy}}>{shift.name}</div>
                        <div style={{fontSize:10,color:C.muted}}>{shift.hours}</div>
                      </div>
                    </div>
                    {mgrNames.length > 0 && (
                      <div style={{marginBottom:8}}>
                        <div style={{fontSize:10,fontWeight:700,color:C.purple,marginBottom:4,textTransform:"uppercase",letterSpacing:.4}}>אחראים ({mgrNames.length}/{shift.maxManagers})</div>
                        {mgrNames.map(function(name,i){
                          return <div key={i} style={{fontSize:12,fontWeight:600,color:"#4C1D95",padding:"2px 0"}}>{name}</div>;
                        })}
                      </div>
                    )}
                    <div>
                      <div style={{fontSize:10,fontWeight:700,color:C.blue,marginBottom:4,textTransform:"uppercase",letterSpacing:.4}}>מתנדבים ({volNames.length}/{shift.maxVolunteers})</div>
                      {volNames.length === 0
                        ? <div style={{fontSize:11,color:C.muted,fontStyle:"italic"}}>אין נרשמים</div>
                        : volNames.map(function(name,i){
                          return <div key={i} style={{fontSize:12,color:C.text,padding:"2px 0",borderBottom:i<volNames.length-1?"1px solid #F0F4F8":"none"}}>{name}</div>;
                        })
                      }
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function AdminPanel(props) {
  var st = useState("shifts"); var tab = st[0]; var setTab = st[1];
  var sf = useState("all"); var filterDay = sf[0]; var setFilterDay = sf[1];
  var ss = useState(""); var search = ss[0]; var setSearch = ss[1];
  var sr = useState(null); var removing = sr[0]; var setRemoving = sr[1];
  var ss2 = useState(null); var selectedShift = ss2[0]; var setSelectedShift = ss2[1];

  var volRegs = 0; var mgrRegs = 0;
  var rkeys = Object.keys(props.regs);
  for (var i = 0; i < rkeys.length; i++) {
    var ut = (props.users[rkeys[i]]||{}).type;
    if (ut==="volunteer") volRegs++;
    if (ut==="manager") mgrRegs++;
  }
  var dmCount = Object.keys(props.dmRegs).length;
  var totalVol = 0;
  for (var j = 0; j < props.shifts.length; j++) totalVol += props.shifts[j].maxVolunteers;
  var fullShifts = 0;
  for (var k = 0; k < props.shifts.length; k++) {
    var s = props.shifts[k];
    var o = props.occ[s.id] || {volunteers:[],managers:[]};
    if (o.volunteers.length >= s.maxVolunteers && o.managers.length >= s.maxManagers) fullShifts++;
  }
  var unfilledMgr = 0;
  for (var m = 0; m < props.shifts.length; m++) {
    if ((props.occ[props.shifts[m].id]||{managers:[]}).managers.length === 0) unfilledMgr++;
  }

  var TABS = [
    ["shifts","📅 משמרות"],["day","👥 לפי יום"],["table","📋 נרשמים"],
    ["allusers","🗂 כל המשתמשים"],["unregistered","⏳ טרם נרשמו"],["log","📜 לוג"],
  ];
  if (props.isSup) {
    TABS.push(["import","⬆ ייבוא"]);
    TABS.push(["config","⚙ הגדרות"]);
  }
  var NO_DAY = ["import","allusers","unregistered","config","log"];

  var regList = [];
  for (var ri = 0; ri < rkeys.length; ri++) {
    var rcode = rkeys[ri];
    var ru = props.users[rcode];
    var rshiftId = props.regs[rcode];
    var rshift = null;
    for (var si = 0; si < props.shifts.length; si++) {
      if (props.shifts[si].id === rshiftId) { rshift = props.shifts[si]; break; }
    }
    if (filterDay !== "all" && (!rshift || rshift.day !== Number(filterDay))) continue;
    if (search) {
      var q = search.toLowerCase();
      var nm = ru ? (ru.name||"") : "";
      if (nm.toLowerCase().indexOf(q)<0 && rcode.indexOf(q)<0 && (ru?ru.phone||"":"").indexOf(q)<0 && (ru?ru.hr||"":"").toLowerCase().indexOf(q)<0) continue;
    }
    regList.push({code:rcode, user:ru, shift:rshift});
  }

  var maxDmCap = 0;
  for (var dc = 1; dc <= 10; dc++) maxDmCap += (props.dayConfigs[dc]||{maxDayMgr:2}).maxDayMgr;

  var stats = [
    {label:"נרשמו בסה\"כ", val:volRegs+mgrRegs+dmCount, icon:"👥", col:C.navy},
    {label:"מתנדבים",       val:volRegs+"/"+totalVol,     icon:"🙋", col:C.blue},
    {label:"אחראי משמרת",  val:mgrRegs+"/"+props.shifts.length, icon:"⭐", col:C.purple},
    {label:"אחראי יום",    val:dmCount+"/"+maxDmCap,      icon:"📋", col:C.teal},
    {label:"משמרות מלאות", val:fullShifts+"/"+props.shifts.length, icon:"✅", col:C.green},
    {label:"ללא אחראי משמרת", val:unfilledMgr,            icon:"⚠", col:C.red},
  ];

  return (
    <div dir="rtl" style={{minHeight:"100vh",background:C.bg,fontFamily:"'Segoe UI',Arial,sans-serif",color:C.text}}>
      <Hdr icon={props.isSup?"👑":"🛠"} title={props.isSup?"פאנל מנהל ראשי":"פאנל מנהל"} name={props.me.name} sub={(props.isSup?"מנהל ראשי":"מנהל")+" - ת.ז. "+props.me.id} onLogout={props.onLogout} bg={props.isSup?"linear-gradient(135deg,#1A1A2E,#16213E)":undefined} />
      <div style={{maxWidth:1120,margin:"0 auto",padding:"24px 16px"}}>

        <div style={{background:props.regOpen?"linear-gradient(135deg,#1D6F42,#27AE60)":"linear-gradient(135deg,#7F1D1D,#E74C3C)",borderRadius:12,padding:"14px 20px",marginBottom:22,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:26}}>{props.regOpen?"🟢":"🔴"}</span>
            <div>
              <div style={{color:"#fff",fontWeight:900,fontSize:16}}>הרשמה - {props.regOpen?"פתוחה":"סגורה"}</div>
              <div style={{color:"rgba(255,255,255,.8)",fontSize:12,marginTop:1}}>{props.regOpen?"מתנדבים יכולים להירשם":"אף אחד אינו יכול להירשם"}</div>
            </div>
          </div>
          {props.isSup && (
            <button onClick={props.onToggleReg} style={{padding:"9px 22px",borderRadius:9,border:"2px solid rgba(255,255,255,.6)",background:"rgba(255,255,255,.15)",color:"#fff",fontSize:13,fontWeight:800,cursor:"pointer"}}>
              {props.regOpen?"🔒 סגור הרשמה":"🔓 פתח הרשמה"}
            </button>
          )}
        </div>

        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))",gap:12,marginBottom:24}}>
          {stats.map(function(st) {
            return (
              <div key={st.label} style={{background:C.card,borderRadius:12,padding:"15px 17px",boxShadow:"0 2px 8px rgba(0,0,0,.07)",borderTop:"4px solid "+st.col}}>
                <div style={{fontSize:19,marginBottom:5}}>{st.icon}</div>
                <div style={{fontSize:22,fontWeight:900,color:st.col,lineHeight:1}}>{st.val}</div>
                <div style={{fontSize:11,color:C.muted,marginTop:3}}>{st.label}</div>
              </div>
            );
          })}
        </div>

        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:18}}>
          <div style={{display:"flex",gap:3,background:"#fff",borderRadius:10,padding:3,boxShadow:"0 1px 4px rgba(0,0,0,.08)",flexWrap:"wrap"}}>
            {TABS.map(function(t) {
              return (
                <button key={t[0]} onClick={function(){setTab(t[0]);}} style={{padding:"6px 11px",borderRadius:7,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:tab===t[0]?C.navy:"transparent",color:tab===t[0]?"#fff":C.muted}}>
                  {t[1]}
                </button>
              );
            })}
          </div>
          <button onClick={function(){doExport(props.users,props.shifts,props.regs,props.dmRegs,props.dayNames);}} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 16px",borderRadius:9,border:"none",background:"linear-gradient(135deg,#1D6F42,#27AE60)",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer"}}>
            ⬇ ייצוא לאקסל
          </button>
        </div>

        {!NO_DAY.includes(tab) && (
          <div style={{display:"flex",gap:7,flexWrap:"wrap",marginBottom:18}}>
            <DayPill label="כל הימים" active={filterDay==="all"} onClick={function(){setFilterDay("all");}} />
            {DAYS.map(function(d) {
              return <DayPill key={d} label={props.dayNames[d]||("יום "+d)} active={filterDay===String(d)} onClick={function(){setFilterDay(String(d));}} />;
            })}
          </div>
        )}

        {tab === "shifts"       && <ShiftsGrid shifts={props.shifts} occ={props.occ} dayNames={props.dayNames} filterDay={filterDay} onShiftClick={function(shift){setSelectedShift(shift);}} />}
        {selectedShift && (
          <AdminRegisterModal
            shift={selectedShift}
            users={props.users}
            regs={props.regs}
            occ={props.occ}
            dayNames={props.dayNames}
            onRegister={function(userId){ props.onAdminRegister(userId, selectedShift.id); setSelectedShift(null); }}
            onClose={function(){ setSelectedShift(null); }}
          />
        )}
        {tab === "day"          && <DayView    shifts={props.shifts} occ={props.occ} dmOcc={props.dmOcc} dayNames={props.dayNames} dayConfigs={props.dayConfigs} users={props.users} filterDay={filterDay} regs={props.dmRegs} onDayClick={function(day){setSelectedShift({isDayMgr:true,day:day});}} />}
        {selectedShift && selectedShift.isDayMgr && (
          <AdminDayMgrModal
            day={selectedShift.day}
            users={props.users}
            dmRegs={props.dmRegs}
            dmOcc={props.dmOcc}
            dayConfigs={props.dayConfigs}
            dayNames={props.dayNames}
            onRegister={function(userId){ props.onAdminDmRegister(userId, selectedShift.day); setSelectedShift(null); }}
            onClose={function(){ setSelectedShift(null); }}
          />
        )}
        {tab === "table"        && <RegTable   regList={regList} search={search} onSearch={setSearch} dayNames={props.dayNames} removing={removing} onRemove={function(c){props.onRemove(c);setRemoving(null);}} onSetRemoving={setRemoving} />}
        {tab === "allusers"     && <AllUsers   users={props.users} shifts={props.shifts} regs={props.regs} dmRegs={props.dmRegs} dayNames={props.dayNames} isSup={props.isSup} onDeleteUser={props.onDeleteUser} onDmRemove={props.onDmRemove} />}
        {tab === "unregistered" && <Unreg      users={props.users} regs={props.regs} dmRegs={props.dmRegs} />}
        {tab === "log"          && <LogView    log={props.log} dayNames={props.dayNames} />}
        {tab === "import"       && props.isSup && <ImportView users={props.users} regs={props.regs} dmRegs={props.dmRegs} shifts={props.shifts} dayNames={props.dayNames} onImport={props.onImport} />}
        {tab === "config"       && props.isSup && <ConfigView shiftMap={props.shiftMap} dayNames={props.dayNames} dayConfigs={props.dayConfigs} occ={props.occ} dmOcc={props.dmOcc} regs={props.regs} onUpdateShift={props.onUpdateShift} onAddShift={props.onAddShift} onRemoveShift={props.onRemoveShift} onUpdateDayName={props.onUpdateDayName} onUpdateDayConfig={props.onUpdateDayConfig} allowSelfRemove={props.allowSelfRemove} onToggleSelfRemove={props.onToggleSelfRemove} />}
      </div>
    </div>
  );
}

// ─── SHIFTS GRID ─────────────────────────────────────────────────────────────
// ─── ADMIN REGISTER MODAL ────────────────────────────────────────────────────
function AdminRegisterModal(props) {
  var ss = useState(""); var search = ss[0]; var setSearch = ss[1];
  var st = useState("volunteer"); var regType = st[0]; var setRegType = st[1];

  var shift = props.shift;
  var occ = props.occ[shift.id] || {volunteers:[], managers:[]};

  // Build list of candidates: not yet registered, correct type
  var registeredIds = {};
  var rkeys = Object.keys(props.regs);
  for (var i = 0; i < rkeys.length; i++) registeredIds[rkeys[i]] = true;

  var candidates = [];
  var ukeys = Object.keys(props.users);
  for (var j = 0; j < ukeys.length; j++) {
    var id = ukeys[j];
    var u = props.users[id];
    if (u.type !== regType) continue;
    if (registeredIds[id]) continue;
    if (search) {
      var q = search.toLowerCase();
      if ((u.name||"").toLowerCase().indexOf(q) < 0 && id.indexOf(q) < 0) continue;
    }
    candidates.push({id:id, u:u});
  }
  candidates.sort(function(a,b){ return (a.u.name||"").localeCompare(b.u.name||"","he"); });

  var volFull = occ.volunteers.length >= shift.maxVolunteers;
  var mgrFull = occ.managers.length  >= shift.maxManagers;
  var isFull  = regType === "volunteer" ? volFull : mgrFull;

  return (
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,.5)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={props.onClose}>
      <div dir="rtl" style={{background:C.card,borderRadius:18,width:"100%",maxWidth:500,maxHeight:"85vh",display:"flex",flexDirection:"column",boxShadow:"0 20px 60px rgba(0,0,0,.4)"}} onClick={function(e){e.stopPropagation();}}>

        {/* Header */}
        <div style={{background:C.navy,borderRadius:"18px 18px 0 0",padding:"16px 20px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div>
            <div style={{color:"#fff",fontWeight:800,fontSize:16}}>{shift.icon} {shift.name}</div>
            <div style={{color:"#93C5FD",fontSize:12,marginTop:2}}>{props.dayNames[shift.day]||("יום "+shift.day)} · {shift.hours}</div>
          </div>
          <button onClick={props.onClose} style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:18,fontWeight:700}}>×</button>
        </div>

        {/* Occupancy summary */}
        <div style={{padding:"12px 20px",borderBottom:"1px solid #EEF2F7",display:"flex",gap:16,background:"#F8FAFC"}}>
          <div style={{fontSize:12,color:C.muted}}>
            מתנדבים: <strong style={{color:volFull?C.red:C.green}}>{occ.volunteers.length}/{shift.maxVolunteers}</strong>
          </div>
          <div style={{fontSize:12,color:C.muted}}>
            אחראים: <strong style={{color:mgrFull?C.red:C.green}}>{occ.managers.length}/{shift.maxManagers}</strong>
          </div>
        </div>

        {/* Type selector */}
        <div style={{padding:"12px 20px",borderBottom:"1px solid #EEF2F7",display:"flex",gap:6}}>
          <button onClick={function(){setRegType("volunteer");}} style={{padding:"6px 16px",borderRadius:18,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,background:regType==="volunteer"?C.blue:"#E2E8F0",color:regType==="volunteer"?"#fff":C.muted}}>
            מתנדבים {volFull?"(מלא)":"("+( shift.maxVolunteers-occ.volunteers.length)+" פנויים)"}
          </button>
          <button onClick={function(){setRegType("manager");}} style={{padding:"6px 16px",borderRadius:18,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,background:regType==="manager"?C.purple:"#E2E8F0",color:regType==="manager"?"#fff":C.muted}}>
            אחראים {mgrFull?"(מלא)":"("+(shift.maxManagers-occ.managers.length)+" פנויים)"}
          </button>
        </div>

        {/* Search */}
        <div style={{padding:"12px 20px",borderBottom:"1px solid #EEF2F7"}}>
          <input value={search} onChange={function(e){setSearch(e.target.value);}} placeholder="חפש שם או ת.ז..."
            autoFocus
            style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"2px solid #CBD5E0",fontSize:13,color:"#1A202C",background:"#fff",outline:"none",boxSizing:"border-box"}} />
        </div>

        {/* Candidate list */}
        <div style={{flex:1,overflowY:"auto",padding:"8px 12px"}}>
          {isFull && (
            <div style={{textAlign:"center",padding:"20px 0",color:C.red,fontSize:13,fontWeight:600}}>המשמרת מלאה עבור סוג זה</div>
          )}
          {!isFull && !candidates.length && (
            <div style={{textAlign:"center",padding:"20px 0",color:C.muted,fontSize:13}}>{search?"לא נמצאו תוצאות":"כל המשתמשים מהסוג הזה כבר רשומים"}</div>
          )}
          {!isFull && candidates.map(function(c) {
            return (
              <div key={c.id} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 8px",borderRadius:9,marginBottom:4,cursor:"pointer",transition:"background 0.1s"}}
                onMouseEnter={function(e){e.currentTarget.style.background="#F0F4F8";}}
                onMouseLeave={function(e){e.currentTarget.style.background="transparent";}}>
                <div style={{width:36,height:36,borderRadius:"50%",background:regType==="manager"?"#EDE9FE":"#DBEAFE",color:regType==="manager"?C.purple:C.blue,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,flexShrink:0}}>
                  {(c.u.name||"?").charAt(0)}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:14,color:C.text}}>{c.u.name}</div>
                  <div style={{display:"flex",gap:10,marginTop:2}}>
                    <span style={{fontSize:11,color:C.muted,fontFamily:"monospace"}}>ת.ז. {c.id}</span>
                    {c.u.phone && <span style={{fontSize:11,color:C.muted}}>{c.u.phone}</span>}
                    {c.u.hr && <span style={{fontSize:11,background:"#F0F4F8",borderRadius:4,padding:"1px 6px",color:C.navy}}>{c.u.hr}</span>}
                  </div>
                </div>
                <button onClick={function(){ props.onRegister(c.id); }}
                  style={{padding:"7px 16px",borderRadius:8,border:"none",background:regType==="manager"?C.purple:C.blue,color:"#fff",fontSize:12,fontWeight:800,cursor:"pointer",flexShrink:0}}>
                  רשום
                </button>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// --- ADMIN DAY MANAGER MODAL -------------------------------------------------
function AdminDayMgrModal(props) {
  var ss = useState(""); var search = ss[0]; var setSearch = ss[1];
  var day = props.day;
  var occupied = (props.dmOcc[day]||[]).length;
  var maxSlots = (props.dayConfigs[day]||{maxDayMgr:2}).maxDayMgr;
  var isFull = occupied >= maxSlots;

  var registeredDmIds = {};
  var dmkeys = Object.keys(props.dmRegs);
  for (var i = 0; i < dmkeys.length; i++) registeredDmIds[dmkeys[i]] = true;

  var candidates = [];
  var ukeys = Object.keys(props.users);
  for (var j = 0; j < ukeys.length; j++) {
    var id = ukeys[j];
    var u = props.users[id];
    if (u.type !== "day_manager") continue;
    if (registeredDmIds[id]) continue;
    if (search) {
      var q = search.toLowerCase();
      if ((u.name||"").toLowerCase().indexOf(q) < 0 && id.indexOf(q) < 0) continue;
    }
    candidates.push({id:id, u:u});
  }
  candidates.sort(function(a,b){ return (a.u.name||"").localeCompare(b.u.name||"","he"); });

  var currentMgrs = (props.dmOcc[day]||[]).map(function(id){ return props.users[id]; }).filter(Boolean);

  return (
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,.5)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={props.onClose}>
      <div dir="rtl" style={{background:C.card,borderRadius:18,width:"100%",maxWidth:480,maxHeight:"80vh",display:"flex",flexDirection:"column",boxShadow:"0 20px 60px rgba(0,0,0,.4)"}} onClick={function(e){e.stopPropagation();}}>

        <div style={{background:C.teal,borderRadius:"18px 18px 0 0",padding:"16px 20px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div>
            <div style={{color:"#fff",fontWeight:800,fontSize:16}}>📋 רישום אחראי יום</div>
            <div style={{color:"rgba(255,255,255,.8)",fontSize:12,marginTop:2}}>{props.dayNames[day]||("יום "+day)}</div>
          </div>
          <button onClick={props.onClose} style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:18,fontWeight:700}}>×</button>
        </div>

        <div style={{padding:"12px 20px",borderBottom:"1px solid #EEF2F7",background:"#F0FDF9"}}>
          <div style={{fontSize:12,color:C.teal,fontWeight:700,marginBottom:6}}>
            נרשמו: {occupied}/{maxSlots} אחראי יום
          </div>
          {currentMgrs.length > 0 && (
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {currentMgrs.map(function(m,i){
                return (
                  <div key={i} style={{display:"flex",alignItems:"center",gap:6,background:"#CCFBF1",borderRadius:7,padding:"4px 10px"}}>
                    <div style={{width:22,height:22,borderRadius:"50%",background:C.teal,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800}}>{(m.name||"?").charAt(0)}</div>
                    <span style={{fontSize:12,fontWeight:700,color:"#134E4A"}}>{m.name}</span>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div style={{padding:"12px 20px",borderBottom:"1px solid #EEF2F7"}}>
          <input value={search} onChange={function(e){setSearch(e.target.value);}} placeholder="חפש שם או ת.ז..." autoFocus
            style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"2px solid #CBD5E0",fontSize:13,color:"#1A202C",background:"#fff",outline:"none",boxSizing:"border-box"}} />
        </div>

        <div style={{flex:1,overflowY:"auto",padding:"8px 12px"}}>
          {isFull && <div style={{textAlign:"center",padding:"20px 0",color:C.teal,fontSize:13,fontWeight:600}}>היום מלא ({maxSlots}/{maxSlots} אחראי יום)</div>}
          {!isFull && !candidates.length && <div style={{textAlign:"center",padding:"20px 0",color:C.muted,fontSize:13}}>{search?"לא נמצאו תוצאות":"כל אחראי היום כבר רשומים"}</div>}
          {!isFull && candidates.map(function(c) {
            return (
              <div key={c.id} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 8px",borderRadius:9,marginBottom:4,cursor:"pointer"}}
                onMouseEnter={function(e){e.currentTarget.style.background="#F0FDF9";}}
                onMouseLeave={function(e){e.currentTarget.style.background="transparent";}}>
                <div style={{width:36,height:36,borderRadius:"50%",background:"#CCFBF1",color:C.teal,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,flexShrink:0}}>
                  {(c.u.name||"?").charAt(0)}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:14,color:C.text}}>{c.u.name}</div>
                  <div style={{display:"flex",gap:10,marginTop:2}}>
                    <span style={{fontSize:11,color:C.muted,fontFamily:"monospace"}}>ת.ז. {c.id}</span>
                    {c.u.phone && <span style={{fontSize:11,color:C.muted}}>{c.u.phone}</span>}
                    {c.u.hr && <span style={{fontSize:11,background:"#F0F4F8",borderRadius:4,padding:"1px 6px",color:C.navy}}>{c.u.hr}</span>}
                  </div>
                </div>
                <button onClick={function(){ props.onRegister(c.id); }}
                  style={{padding:"7px 16px",borderRadius:8,border:"none",background:C.teal,color:"#fff",fontSize:12,fontWeight:800,cursor:"pointer",flexShrink:0}}>
                  רשום
                </button>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function ShiftsGrid(props) {
  var visibleDays = props.filterDay === "all" ? DAYS : [Number(props.filterDay)];
  var clickable = !!props.onShiftClick;
  return (
    <div>
      {clickable && <div style={{fontSize:12,color:C.muted,marginBottom:14}}>לחץ על משמרת כדי לרשום אליה מתנדב/ת</div>}
      {visibleDays.map(function(day) {
        var dayShifts = props.shifts.filter(function(s){ return s.day === day; });
        if (!dayShifts.length) return null;
        return (
          <div key={day} style={{marginBottom:24}}>
            <div style={{marginBottom:10}}><span style={{background:C.navy,color:"#fff",borderRadius:6,padding:"2px 11px",fontSize:12,fontWeight:700}}>{props.dayNames[day]||("יום "+day)}</span></div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:12}}>
              {dayShifts.map(function(shift) {
                var o = props.occ[shift.id] || {volunteers:[],managers:[]};
                var vPct = shift.maxVolunteers > 0 ? o.volunteers.length / shift.maxVolunteers : 0;
                var mOk = o.managers.length >= shift.maxManagers;
                return (
                  <div key={shift.id}
                    onClick={clickable ? function(){ props.onShiftClick(shift); } : undefined}
                    style={{background:C.card,borderRadius:13,padding:16,boxShadow:"0 2px 8px rgba(0,0,0,.08)",cursor:clickable?"pointer":"default",transition:clickable?"box-shadow 0.15s":"none",border:"2px solid transparent"}}
                    onMouseEnter={clickable?function(e){e.currentTarget.style.boxShadow="0 4px 16px rgba(15,45,74,.2)";e.currentTarget.style.borderColor="#CBD5E0";}:undefined}
                    onMouseLeave={clickable?function(e){e.currentTarget.style.boxShadow="0 2px 8px rgba(0,0,0,.08)";e.currentTarget.style.borderColor="transparent";}:undefined}
                  >
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                      <div style={{display:"flex",alignItems:"center",gap:7}}>
                        <span style={{fontSize:18}}>{shift.icon}</span>
                        <span style={{fontSize:15,fontWeight:800,color:C.navy}}>{shift.name}</span>
                      </div>
                      <span style={{fontSize:10,color:C.muted}}>{shift.hours}</span>
                    </div>
                    <div style={{marginBottom:8}}>
                      <Bar val={o.volunteers.length} max={shift.maxVolunteers} label="מתנדבים" />
                    </div>
                    <div style={{background:mOk?"#D5F5E3":"#FEF9E7",borderRadius:7,padding:"5px 10px",display:"flex",justifyContent:"space-between"}}>
                      <span style={{fontSize:11,color:C.muted}}>אחראי משמרת</span>
                      <span style={{fontSize:11,fontWeight:800,color:mOk?C.green:C.amber}}>{o.managers.length}/{shift.maxManagers} {mOk?"מאויש":"ממתין"}</span>
                    </div>
                    {clickable && (
                      <div style={{marginTop:10,textAlign:"center",fontSize:11,color:C.blue,fontWeight:600}}>+ רשום מתנדב/ת</div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ─── DAY VIEW ─────────────────────────────────────────────────────────────────
function DayView(props) {
  var visibleDays = props.filterDay === "all" ? DAYS : [Number(props.filterDay)];
  return (
    <div>
      {visibleDays.map(function(day) {
        var dayShifts = props.shifts.filter(function(s){ return s.day === day; });
        if (!dayShifts.length) return null;
        var totalReg = 0;
        for (var i = 0; i < dayShifts.length; i++) {
          var o = props.occ[dayShifts[i].id]||{volunteers:[],managers:[]};
          totalReg += o.volunteers.length + o.managers.length;
        }
        totalReg += (props.dmOcc[day]||[]).length;
        var dayMgrs = (props.dmOcc[day]||[]).map(function(id){ return props.users[id]; }).filter(Boolean);
        var maxDm = (props.dayConfigs[day]||{maxDayMgr:2}).maxDayMgr;
        return (
          <div key={day} style={{background:C.card,borderRadius:16,marginBottom:20,boxShadow:"0 2px 10px rgba(0,0,0,.08)",overflow:"hidden"}}>
            <div style={{background:C.navy,color:"#fff",padding:"12px 20px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:18,fontWeight:900}}>{props.dayNames[day]||("יום "+day)}</span>
                <span style={{fontSize:12,opacity:.7}}>{totalReg} נרשמים - {dayShifts.length} משמרות</span>
              </div>
              {props.onDayClick && (
                <button onClick={function(){props.onDayClick(day);}}
                  style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.4)",color:"#fff",borderRadius:8,padding:"5px 14px",cursor:"pointer",fontSize:12,fontWeight:700}}>
                  + רשום אחראי יום
                </button>
              )}
            </div>
            <div style={{padding:"12px 18px",borderBottom:"2px solid #EEF2F7",background:"#F0FDF9"}}>
              <div style={{fontSize:10,fontWeight:700,color:C.teal,marginBottom:7,textTransform:"uppercase",letterSpacing:.5}}>אחראי יום ({dayMgrs.length}/{maxDm})</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {dayMgrs.length === 0 && <span style={{fontSize:11,color:C.muted,fontStyle:"italic"}}>אין אחראי יום</span>}
                {dayMgrs.map(function(m,i) {
                  return (
                    <div key={i} style={{display:"flex",alignItems:"center",gap:7,background:"#CCFBF1",borderRadius:7,padding:"5px 10px"}}>
                      <div style={{width:24,height:24,borderRadius:"50%",background:C.teal,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800}}>{(m.name||"?").charAt(0)}</div>
                      <div>
                        <div style={{fontSize:12,fontWeight:700,color:"#134E4A"}}>{m.name}</div>
                        <div style={{fontSize:10,color:C.muted}}>{m.phone}</div>
                      </div>
                    </div>
                  );
                })}
                {dayMgrs.length < maxDm && Array(maxDm - dayMgrs.length).fill(0).map(function(_,i) {
                  return <div key={"e"+i} style={{display:"flex",alignItems:"center",gap:5,border:"1.5px dashed #5EEAD4",borderRadius:7,padding:"5px 10px"}}><span style={{fontSize:11,color:"#5EEAD4",fontWeight:600}}>+ ממתין</span></div>;
                })}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat("+Math.min(dayShifts.length,4)+",1fr)"}}>
              {dayShifts.map(function(shift, idx) {
                var o = props.occ[shift.id]||{volunteers:[],managers:[]};
                var mgrs = o.managers.map(function(id){ return props.users[id]; }).filter(Boolean);
                var vols = o.volunteers.map(function(id){ return props.users[id]; }).filter(Boolean);
                return (
                  <div key={shift.id} style={{padding:"16px 18px",borderLeft:idx>0?"1px solid #EEF2F7":"none",minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:12,paddingBottom:10,borderBottom:"2px solid #EEF2F7"}}>
                      <span style={{fontSize:16}}>{shift.icon}</span>
                      <div>
                        <div style={{fontSize:13,fontWeight:800,color:C.navy}}>{shift.name}</div>
                        <div style={{fontSize:10,color:C.muted}}>{shift.hours}</div>
                      </div>
                    </div>
                    <div style={{marginBottom:8}}>
                      <div style={{fontSize:10,fontWeight:700,color:C.muted,marginBottom:4,textTransform:"uppercase"}}>אחראי ({mgrs.length}/{shift.maxManagers})</div>
                      {mgrs.length === 0
                        ? <div style={{background:"#FEF9E7",border:"1.5px dashed #F39C12",borderRadius:7,padding:"5px 9px",fontSize:11,color:C.amber,fontWeight:600,textAlign:"center"}}>ממתין</div>
                        : mgrs.map(function(mgr,mi){
                          return (
                            <div key={mi} style={{display:"flex",alignItems:"center",gap:6,background:"#EDE9FE",borderRadius:7,padding:"5px 9px",marginBottom:mi<mgrs.length-1?4:0}}>
                              <div style={{width:22,height:22,borderRadius:"50%",background:C.purple,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,flexShrink:0}}>{(mgr.name||"?").charAt(0)}</div>
                              <div><div style={{fontSize:11,fontWeight:700,color:"#4C1D95"}}>{mgr.name}</div><div style={{fontSize:10,color:C.muted}}>{mgr.phone}</div></div>
                            </div>
                          );
                        })
                      }
                    </div>
                    <div>
                      <div style={{fontSize:10,fontWeight:700,color:C.muted,marginBottom:4,textTransform:"uppercase"}}>מתנדבים ({vols.length}/{shift.maxVolunteers})</div>
                      {vols.length === 0
                        ? <div style={{fontSize:10,color:C.muted,fontStyle:"italic"}}>אין נרשמים</div>
                        : vols.map(function(v,vi) {
                          return (
                            <div key={vi} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 6px",borderRadius:6,background:vi%2===0?"#F8FAFC":"#fff"}}>
                              <span style={{width:20,height:20,borderRadius:"50%",background:C.blue,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,flexShrink:0}}>{(v.name||"?").charAt(0)}</span>
                              <div>
                                <div style={{fontSize:11,fontWeight:700}}>{v.name}</div>
                                <div style={{fontSize:9,color:C.muted}}>{v.phone}</div>
                              </div>
                            </div>
                          );
                        })
                      }
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ─── REGISTRATION TABLE ───────────────────────────────────────────────────────
function RegTable(props) {
  return (
    <div>
      <div style={{marginBottom:14,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
        <input value={props.search} onChange={function(e){props.onSearch(e.target.value);}} placeholder="חפש שם, ת.ז., טלפון, HR..."
          style={{padding:"9px 14px",borderRadius:9,border:"2px solid #CBD5E0",fontSize:13,color:"#1A202C",background:"#fff",outline:"none",width:"100%",maxWidth:340,boxSizing:"border-box"}} />
        <span style={{fontSize:12,color:C.muted}}>{props.regList.length} רשומות</span>
      </div>
      {!props.regList.length
        ? <div style={{textAlign:"center",padding:60,color:C.muted,background:C.card,borderRadius:14,boxShadow:"0 2px 8px rgba(0,0,0,.07)"}}><div style={{fontSize:38,marginBottom:10}}>📭</div><div>אין רישומים</div></div>
        : (
          <div style={{background:C.card,borderRadius:14,overflow:"hidden",boxShadow:"0 2px 10px rgba(0,0,0,.08)"}}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:720}}>
                <thead>
                  <tr style={{background:C.navy,color:"#fff"}}>
                    {["שם","ת.ז.","סוג","טלפון","אימייל","HR","יום","משמרת","שעות",""].map(function(h,i) {
                      return <th key={i} style={{padding:"11px 13px",textAlign:"right",fontWeight:700,fontSize:11,whiteSpace:"nowrap"}}>{h}</th>;
                    })}
                  </tr>
                </thead>
                <tbody>
                  {props.regList.map(function(r,i) {
                    var utype = r.user ? r.user.type : "";
                    var hr = r.user ? (r.user.hr||"") : "";
                    return (
                      <tr key={r.code} style={{background:i%2===0?"#fff":"#F8FAFC",borderBottom:"1px solid #EEF2F7"}}>
                        <td style={{padding:"10px 13px",fontWeight:700}}>{r.user ? r.user.name : "-"}</td>
                        <td style={{padding:"10px 13px",fontFamily:"monospace",color:C.muted,fontSize:11}}>{r.code}</td>
                        <td style={{padding:"10px 13px"}}>
                          <span style={{background:utype==="manager"?"#EDE9FE":utype==="day_manager"?"#CCFBF1":"#DBEAFE",color:utype==="manager"?C.purple:utype==="day_manager"?C.teal:C.blue,borderRadius:5,padding:"2px 8px",fontSize:11,fontWeight:700}}>
                            {utype==="manager"?"אחראי משמרת":utype==="day_manager"?"אחראי יום":"מתנדב"}
                          </span>
                        </td>
                        <td style={{padding:"10px 13px",color:C.muted}}>{r.user ? r.user.phone||"-" : "-"}</td>
                        <td style={{padding:"10px 13px",color:C.muted}}>{r.user ? r.user.email||"-" : "-"}</td>
                        <td style={{padding:"10px 13px",maxWidth:160}}>
                          {hr
                            ? <span style={{background:"#F0F4F8",borderRadius:5,padding:"2px 8px",fontSize:11,color:C.navy}}>{hr}</span>
                            : <span style={{color:"#CBD5E0",fontSize:11}}>-</span>
                          }
                        </td>
                        <td style={{padding:"10px 13px",fontWeight:700,color:C.navy}}>{r.shift ? (props.dayNames[r.shift.day]||("יום "+r.shift.day)) : "-"}</td>
                        <td style={{padding:"10px 13px"}}>{r.shift ? (r.shift.icon+" "+r.shift.name) : "-"}</td>
                        <td style={{padding:"10px 13px",color:C.muted,fontSize:11}}>{r.shift ? r.shift.hours : "-"}</td>
                        <td style={{padding:"10px 13px"}}>
                          {props.removing === r.code
                            ? <div style={{display:"flex",gap:4}}><button onClick={function(){props.onRemove(r.code);}} style={{background:C.red,color:"#fff",border:"none",borderRadius:5,padding:"3px 9px",fontSize:11,fontWeight:700,cursor:"pointer"}}>אשר</button><button onClick={function(){props.onSetRemoving(null);}} style={{background:"#E2E8F0",border:"none",borderRadius:5,padding:"3px 8px",fontSize:11,cursor:"pointer"}}>ביטול</button></div>
                            : <button onClick={function(){props.onSetRemoving(r.code);}} style={{background:"#FEF2F2",border:"1px solid "+C.red,color:C.red,borderRadius:6,padding:"4px 11px",fontSize:11,fontWeight:700,cursor:"pointer"}}>הסר</button>
                          }
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )
      }
    </div>
  );
}

// ─── ALL USERS ────────────────────────────────────────────────────────────────
function AllUsers(props) {
  var ss = useState(""); var search = ss[0]; var setSearch = ss[1];
  var sf = useState("all"); var ft = sf[0]; var setFt = sf[1];
  var sd = useState(null); var delConfirm = sd[0]; var setDelConfirm = sd[1];
  var sr = useState(null); var removingDm = sr[0]; var setRemovingDm = sr[1];

  var allList = [];
  var ukeys = Object.keys(props.users);
  for (var i = 0; i < ukeys.length; i++) {
    var id = ukeys[i];
    var u = props.users[id];
    if (!props.isSup && (u.type === "admin" || u.type === "superadmin")) continue;
    if (ft !== "all" && u.type !== ft) continue;
    if (search) {
      var q = search.toLowerCase();
      if ((u.name||"").toLowerCase().indexOf(q)<0 && id.indexOf(q)<0 && (u.phone||"").indexOf(q)<0 && (u.email||"").toLowerCase().indexOf(q)<0 && (u.hr||"").toLowerCase().indexOf(q)<0) continue;
    }
    var shiftId = props.regs[id] || null;
    var shift = null;
    if (shiftId) {
      for (var j = 0; j < props.shifts.length; j++) {
        if (props.shifts[j].id === shiftId) { shift = props.shifts[j]; break; }
      }
    }
    var dayNum = props.dmRegs[id] || null;
    allList.push({id:id, u:u, shift:shift, shiftId:shiftId, dayNum:dayNum});
  }
  var ORDER = {superadmin:0,admin:1,day_manager:2,manager:3,volunteer:4};
  allList.sort(function(a,b){
    var od = (ORDER[a.u.type]||9) - (ORDER[b.u.type]||9);
    if (od!==0) return od;
    return (a.u.name||"").localeCompare(b.u.name||"","he");
  });

  var pills = [["all","הכל"],["volunteer","מתנדבים"],["manager","אחראי משמרת"],["day_manager","אחראי יום"]];
  if (props.isSup) pills.push(["admin","מנהלים"]);

  var hdrs = ["שם","ת.ז.","סוג","טלפון","אימייל","סטטוס","שיבוץ","HR"];
  if (props.isSup) { hdrs.push("סיסמה"); hdrs.push(""); }

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:12}}>
        <input value={search} onChange={function(e){setSearch(e.target.value);}} placeholder="חפש..."
          style={{padding:"9px 14px",borderRadius:9,border:"2px solid #CBD5E0",fontSize:13,color:"#1A202C",background:"#fff",outline:"none",width:"100%",maxWidth:260,boxSizing:"border-box"}} />
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
          {pills.map(function(p) {
            return <button key={p[0]} onClick={function(){setFt(p[0]);}} style={{padding:"5px 12px",borderRadius:18,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:ft===p[0]?C.navy:"#fff",color:ft===p[0]?"#fff":C.navy,boxShadow:ft===p[0]?"0 2px 8px rgba(15,45,74,.3)":"0 1px 3px rgba(0,0,0,.1)"}}>{p[1]}</button>;
          })}
        </div>
      </div>
      <div style={{background:C.card,borderRadius:14,overflow:"hidden",boxShadow:"0 2px 10px rgba(0,0,0,.08)"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:600}}>
            <thead>
              <tr style={{background:C.navy,color:"#fff"}}>
                {hdrs.map(function(h,i){ return <th key={i} style={{padding:"10px 12px",textAlign:"right",fontWeight:700,fontSize:11,whiteSpace:"nowrap"}}>{h}</th>; })}
              </tr>
            </thead>
            <tbody>
              {allList.map(function(row,i) {
                var tm = TYPE_INFO[row.u.type] || {label:row.u.type,bg:"#F1F5F9",col:C.muted};
                var isReg = !!(row.shiftId || row.dayNum);
                var isProtected = row.u.type === "superadmin";
                var assignText = "-";
                if (row.shift) assignText = row.shift.icon+" "+(props.dayNames[row.shift.day]||("יום "+row.shift.day))+" - "+row.shift.name;
                else if (row.dayNum) assignText = "📋 "+(props.dayNames[row.dayNum]||("יום "+row.dayNum));
                return (
                  <tr key={row.id} style={{background:i%2===0?"#fff":"#F8FAFC",borderBottom:"1px solid #EEF2F7"}}>
                    <td style={{padding:"9px 12px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:7}}>
                        <div style={{width:26,height:26,borderRadius:"50%",background:tm.bg,color:tm.col,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800,flexShrink:0}}>{(row.u.name||"?").charAt(0)}</div>
                        <span style={{fontWeight:700}}>{row.u.name}</span>
                      </div>
                    </td>
                    <td style={{padding:"9px 12px",fontFamily:"monospace",color:C.muted,fontSize:10}}>{row.id}</td>
                    <td style={{padding:"9px 12px"}}><span style={{background:tm.bg,color:tm.col,borderRadius:5,padding:"2px 7px",fontSize:10,fontWeight:700}}>{tm.label}</span></td>
                    <td style={{padding:"9px 12px",color:C.muted}}>{row.u.phone||"-"}</td>
                    <td style={{padding:"9px 12px",color:C.muted}}>{row.u.email||"-"}</td>
                    <td style={{padding:"9px 12px"}}>
                      {row.u.type==="day_manager"
                        ? (row.dayNum
                            ? (removingDm===row.id
                                ? <div style={{display:"flex",gap:4}}><button onClick={function(){props.onDmRemove(row.id);setRemovingDm(null);}} style={{background:C.red,color:"#fff",border:"none",borderRadius:5,padding:"2px 8px",fontSize:10,fontWeight:700,cursor:"pointer"}}>אשר</button><button onClick={function(){setRemovingDm(null);}} style={{background:"#E2E8F0",border:"none",borderRadius:5,padding:"2px 7px",fontSize:10,cursor:"pointer"}}>ביטול</button></div>
                                : <button onClick={function(){setRemovingDm(row.id);}} style={{background:"#FEF2F2",border:"1px solid "+C.red,color:C.red,borderRadius:6,padding:"3px 9px",fontSize:10,fontWeight:700,cursor:"pointer"}}>הסר מיום</button>
                              )
                            : <span style={{background:"#FEF2F2",color:C.red,borderRadius:6,padding:"2px 7px",fontSize:10,fontWeight:700}}>לא רשום</span>
                          )
                        : row.u.type==="admin"||row.u.type==="superadmin"
                          ? <span style={{color:C.muted,fontSize:10}}>-</span>
                          : isReg
                            ? <span style={{background:"#D5F5E3",color:C.green,borderRadius:6,padding:"2px 7px",fontSize:10,fontWeight:700}}>נרשם</span>
                            : <span style={{background:"#FEF2F2",color:C.red,borderRadius:6,padding:"2px 7px",fontSize:10,fontWeight:700}}>לא נרשם</span>
                      }
                    </td>
                    <td style={{padding:"9px 12px",fontSize:11,color:C.text}}>{assignText}</td>
                    <td style={{padding:"9px 12px",fontSize:11,color:C.muted,maxWidth:180}}>
                      {row.u.hr
                        ? <span style={{display:"inline-block",background:"#F0F4F8",borderRadius:5,padding:"2px 8px",fontSize:11,color:C.navy}}>{row.u.hr}</span>
                        : <span style={{color:"#CBD5E0"}}>-</span>
                      }
                    </td>
                    {props.isSup && (
                      <td style={{padding:"9px 12px"}}>
                        {row.u.type==="admin"
                          ? <span style={{fontFamily:"monospace",fontSize:10,background:"#F0F4F8",padding:"2px 6px",borderRadius:4}}>{row.u.password||"-"}</span>
                          : row.u.type==="superadmin"
                            ? <span style={{fontSize:10,color:C.muted}}>מוגנת</span>
                            : <span style={{fontSize:10,color:C.muted}}>-</span>
                        }
                      </td>
                    )}
                    {props.isSup && (
                      <td style={{padding:"9px 12px"}}>
                        {delConfirm===row.id
                          ? <div style={{display:"flex",gap:4}}><button onClick={function(){props.onDeleteUser(row.id);setDelConfirm(null);}} style={{background:C.red,color:"#fff",border:"none",borderRadius:5,padding:"2px 8px",fontSize:10,fontWeight:700,cursor:"pointer"}}>מחק</button><button onClick={function(){setDelConfirm(null);}} style={{background:"#E2E8F0",border:"none",borderRadius:5,padding:"2px 7px",fontSize:10,cursor:"pointer"}}>ביטול</button></div>
                          : isProtected
                            ? <span style={{fontSize:10,color:C.muted}}>-</span>
                            : (row.u.type==="day_manager"?row.dayNum:isReg)
                              ? <span style={{fontSize:10,color:C.muted,fontStyle:"italic"}}>הסר תחילה</span>
                              : <button onClick={function(){setDelConfirm(row.id);}} style={{background:"#FEF2F2",border:"1px solid "+C.red,color:C.red,borderRadius:6,padding:"3px 9px",fontSize:10,fontWeight:700,cursor:"pointer"}}>מחק</button>
                        }
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        {!allList.length && <div style={{textAlign:"center",padding:40,color:C.muted}}><div style={{fontSize:32,marginBottom:8}}>🔍</div><div>לא נמצאו</div></div>}
      </div>
    </div>
  );
}

// ─── UNREGISTERED ─────────────────────────────────────────────────────────────
function Unreg(props) {
  var sf = useState("all"); var ft = sf[0]; var setFt = sf[1];
  var ss = useState(""); var search = ss[0]; var setSearch = ss[1];

  var regIds  = {};
  var rkeys = Object.keys(props.regs);
  for (var i=0;i<rkeys.length;i++) regIds[rkeys[i]] = true;
  var dmIds = {};
  var dkeys = Object.keys(props.dmRegs);
  for (var j=0;j<dkeys.length;j++) dmIds[dkeys[j]] = true;

  var list = [];
  var ukeys = Object.keys(props.users);
  for (var k=0;k<ukeys.length;k++) {
    var id = ukeys[k];
    var u = props.users[id];
    if (u.type==="admin"||u.type==="superadmin") continue;
    var unreg = u.type==="day_manager" ? !dmIds[id] : !regIds[id];
    if (!unreg) continue;
    if (ft!=="all" && u.type!==ft) continue;
    if (search) {
      var q = search.toLowerCase();
      if ((u.name||"").toLowerCase().indexOf(q)<0 && id.indexOf(q)<0 && (u.phone||"").indexOf(q)<0 && (u.hr||"").toLowerCase().indexOf(q)<0) continue;
    }
    list.push({id:id,u:u});
  }
  var ORDER = {manager:0,day_manager:1,volunteer:2};
  list.sort(function(a,b){
    var od=(ORDER[a.u.type]||9)-(ORDER[b.u.type]||9);
    if(od!==0)return od;
    return (a.u.name||"").localeCompare(b.u.name||"","he");
  });

  var tv=0,tm=0,td=0,uv=0,um=0,ud=0;
  for(var ui=0;ui<ukeys.length;ui++){
    var uu=props.users[ukeys[ui]];
    if(uu.type==="volunteer")tv++;
    if(uu.type==="manager")tm++;
    if(uu.type==="day_manager")td++;
  }
  for(var li=0;li<list.length;li++){
    if(list[li].u.type==="volunteer")uv++;
    if(list[li].u.type==="manager")um++;
    if(list[li].u.type==="day_manager")ud++;
  }

  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:12,marginBottom:18}}>
        {[[uv,tv,"מתנדבים שטרם נרשמו",C.blue,"#EFF6FF"],[um,tm,"אחראי משמרת שטרם נרשמו",C.purple,"#F5F3FF"],[ud,td,"אחראי יום שטרם נרשמו",C.teal,"#F0FDF9"]].map(function(s,i){
          var pct = s[1]>0 ? ((s[1]-s[0])/s[1]) : 0;
          return (
            <div key={i} style={{background:s[4],borderRadius:12,padding:"14px 18px",border:"2px solid "+s[3]+"22"}}>
              <div style={{fontSize:24,fontWeight:900,color:s[3]}}>{s[0]}</div>
              <div style={{fontSize:11,color:s[3],fontWeight:600,marginTop:2}}>{s[2]}</div>
              <div style={{marginTop:7,height:5,background:"rgba(0,0,0,.08)",borderRadius:3,overflow:"hidden"}}>
                <div style={{height:"100%",borderRadius:3,width:(pct*100)+"%",background:s[3]}} />
              </div>
            </div>
          );
        })}
      </div>

      <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:14}}>
        <input value={search} onChange={function(e){setSearch(e.target.value);}} placeholder="חפש שם, ת.ז., טלפון, HR..."
          style={{padding:"9px 14px",borderRadius:9,border:"2px solid #CBD5E0",fontSize:13,color:"#1A202C",background:"#fff",outline:"none",width:"100%",maxWidth:300,boxSizing:"border-box"}} />
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          {[["all","הכל ("+(uv+um+ud)+")"],["volunteer","מתנדבים"],["manager","אחראי משמרת"],["day_manager","אחראי יום"]].map(function(p){
            return <button key={p[0]} onClick={function(){setFt(p[0]);}} style={{padding:"5px 13px",borderRadius:18,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:ft===p[0]?C.navy:"#fff",color:ft===p[0]?"#fff":C.navy,boxShadow:ft===p[0]?"0 2px 8px rgba(15,45,74,.3)":"0 1px 3px rgba(0,0,0,.1)"}}>{p[1]}</button>;
          })}
        </div>
        <span style={{fontSize:12,color:C.muted}}>{list.length} רשומות</span>
      </div>

      {!list.length
        ? <div style={{background:C.card,borderRadius:14,padding:50,textAlign:"center"}}>
            {!search && ft==="all"
              ? <><div style={{fontSize:40,marginBottom:10}}>🎉</div><div style={{fontSize:16,fontWeight:800,color:C.green}}>כולם נרשמו!</div></>
              : <><div style={{fontSize:36,marginBottom:10}}>🔍</div><div style={{fontSize:14,color:C.muted}}>לא נמצאו תוצאות</div></>
            }
          </div>
        : <div style={{display:"flex",flexDirection:"column",gap:7}}>
          {list.map(function(row){
            var tm2 = TYPE_INFO[row.u.type] || {label:row.u.type,bg:"#F1F5F9",col:C.muted};
            return (
              <div key={row.id} style={{background:C.card,borderRadius:10,padding:"12px 18px",display:"flex",alignItems:"center",gap:12,boxShadow:"0 1px 5px rgba(0,0,0,.07)",borderRight:"4px solid "+tm2.col}}>
                <div style={{width:36,height:36,borderRadius:"50%",background:tm2.bg,color:tm2.col,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,flexShrink:0}}>{(row.u.name||"?").charAt(0)}</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:14}}>{row.u.name}</div>
                  <div style={{display:"flex",gap:10,marginTop:2,flexWrap:"wrap",alignItems:"center"}}>
                    <span style={{fontSize:11,color:C.muted,fontFamily:"monospace"}}>ת.ז. {row.id}</span>
                    {row.u.phone && <span style={{fontSize:11,color:C.muted}}>📞 {row.u.phone}</span>}
                    {row.u.hr && <span style={{fontSize:11,background:"#F0F4F8",borderRadius:4,padding:"1px 7px",color:C.navy}}>HR: {row.u.hr}</span>}
                  </div>
                </div>
                <span style={{background:tm2.bg,color:tm2.col,borderRadius:7,padding:"3px 9px",fontSize:10,fontWeight:700,flexShrink:0}}>{tm2.label}</span>
                <span style={{background:"#FEF2F2",color:C.red,borderRadius:7,padding:"3px 9px",fontSize:10,fontWeight:700,flexShrink:0}}>טרם נרשם/ה</span>
              </div>
            );
          })}
        </div>
      }
    </div>
  );
}

// ─── LOG ──────────────────────────────────────────────────────────────────────
function LogView(props) {
  var sf = useState("all"); var ft = sf[0]; var setFt = sf[1];
  var ss = useState(""); var search = ss[0]; var setSearch = ss[1];
  var filtered = props.log.filter(function(e){
    if(ft!=="all"&&e.type!==ft) return false;
    if(search){var q=search.toLowerCase();return (e.userName||"").toLowerCase().indexOf(q)>=0||(e.userId||"").indexOf(q)>=0||(e.actorName||"").toLowerCase().indexOf(q)>=0;}
    return true;
  });
  var TM = {
    register:      {label:"רישום",              col:C.green,  bg:"#D5F5E3"},
    remove:        {label:"הסרה",               col:C.red,    bg:"#FEF2F2"},
    delete_user:   {label:"מחיקת משתמש",        col:C.amber,  bg:"#FEF3C7"},
    import_new:    {label:"ייבוא",              col:C.teal,   bg:"#CCFBF1"},
    import_update: {label:"ייבוא - עדכון פרטים",col:C.blue,   bg:"#DBEAFE"},
    import_role:   {label:"ייבוא - עדכון תפקיד",col:C.purple, bg:"#EDE9FE"},
  };
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:14}}>
        <input value={search} onChange={function(e){setSearch(e.target.value);}} placeholder="חפש..."
          style={{padding:"9px 14px",borderRadius:9,border:"2px solid #CBD5E0",fontSize:13,color:"#1A202C",background:"#fff",outline:"none",width:"100%",maxWidth:250,boxSizing:"border-box"}} />
        <div style={{display:"flex",gap:5}}>
          {[["all","הכל"],["register","רישומים"],["remove","הסרות"],["delete_user","מחיקות"],["import_new","ייבוא"],["import_update","עדכון פרטים"],["import_role","עדכון תפקיד"]].map(function(p){
            return <button key={p[0]} onClick={function(){setFt(p[0]);}} style={{padding:"5px 11px",borderRadius:18,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:ft===p[0]?C.navy:"#fff",color:ft===p[0]?"#fff":C.muted,boxShadow:ft===p[0]?"0 2px 8px rgba(15,45,74,.3)":"0 1px 3px rgba(0,0,0,.1)"}}>{p[1]}</button>;
          })}
        </div>
        <button onClick={function(){doExportLog(props.log,props.dayNames);}} style={{padding:"7px 14px",borderRadius:9,border:"none",background:"linear-gradient(135deg,#1D6F42,#27AE60)",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>⬇ ייצוא</button>
      </div>
      {!filtered.length
        ? <div style={{background:C.card,borderRadius:14,padding:50,textAlign:"center",color:C.muted}}><div style={{fontSize:38,marginBottom:10}}>📋</div><div>{!props.log.length?"הלוג ריק":"אין תוצאות"}</div></div>
        : <div style={{background:C.card,borderRadius:14,overflow:"hidden",boxShadow:"0 2px 10px rgba(0,0,0,.08)"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:700}}>
              <thead><tr style={{background:C.navy,color:"#fff"}}>
                {["זמן","פעולה","משתמש","ת.ז.","יום","משמרת","בוצע ע\"י","תפקיד"].map(function(h,i){
                  return <th key={i} style={{padding:"10px 12px",textAlign:"right",fontWeight:700,fontSize:11,whiteSpace:"nowrap"}}>{h}</th>;
                })}
              </tr></thead>
              <tbody>
                {filtered.map(function(e,i){
                  var m = TM[e.type] || {label:e.type,col:C.muted,bg:"#F1F5F9"};
                  return (
                    <tr key={e.id} style={{background:i%2===0?"#fff":"#F8FAFC",borderBottom:"1px solid #EEF2F7"}}>
                      <td style={{padding:"9px 12px",color:C.muted,fontSize:10,whiteSpace:"nowrap"}}>{fmtTime(e.ts)}</td>
                      <td style={{padding:"9px 12px"}}><span style={{display:"inline-flex",alignItems:"center",gap:4,background:m.bg,color:m.col,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>{m.label}</span></td>
                      <td style={{padding:"9px 12px",fontWeight:700}}>{e.userName}</td>
                      <td style={{padding:"9px 12px",fontFamily:"monospace",color:C.muted,fontSize:10}}>{e.userId}</td>
                      <td style={{padding:"9px 12px",color:C.navy,fontWeight:600}}>{e.dayLabel||"-"}</td>
                      <td style={{padding:"9px 12px"}}>{e.shiftName||"-"}</td>
                      <td style={{padding:"9px 12px",fontWeight:600}}>{e.actorName}</td>
                      <td style={{padding:"9px 12px",color:C.muted,fontSize:11}}>{e.actorType}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      }
    </div>
  );
}

// ─── CONFIG (superadmin) ─────────────────────────────────────────────────────
function ConfigView(props) {
  var si = useState(null); var iconPicker = si[0]; var setIconPicker = si[1];
  var sc = useState(null); var confirmDel = sc[0]; var setConfirmDel = sc[1];

  function regCount(shiftId) {
    var keys = Object.keys(props.regs);
    var cnt = 0;
    for(var i=0;i<keys.length;i++) if(props.regs[keys[i]]===shiftId) cnt++;
    return cnt;
  }

  return (
    <div>
      <div style={{background:C.card,borderRadius:14,padding:"16px 22px",marginBottom:16,boxShadow:"0 2px 8px rgba(0,0,0,.08)"}}>
        <h3 style={{color:C.navy,margin:"0 0 4px",fontSize:17,fontWeight:800}}>הגדרת ימים ומשמרות</h3>
        <p style={{color:C.muted,margin:0,fontSize:12}}>ערוך שם יום, אחראי יום, שם/שעות/אייקון/קיבולת לכל משמרת. הוסף או מחק משמרות.</p>
      </div>

      {/* Self-remove toggle */}
      <div style={{background:C.card,borderRadius:14,padding:"16px 22px",marginBottom:20,boxShadow:"0 2px 8px rgba(0,0,0,.08)",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
        <div>
          <div style={{fontSize:15,fontWeight:800,color:C.navy,marginBottom:3}}>אפשרות ביטול עצמי</div>
          <div style={{fontSize:12,color:C.muted}}>כאשר מופעל — מתנדבים ואחראים יכולים לבטל את רישומם ולהירשם מחדש בעצמם.</div>
        </div>
        <button onClick={props.onToggleSelfRemove} style={{
          padding:"10px 22px", borderRadius:10, border:"none", cursor:"pointer",
          fontSize:14, fontWeight:800,
          background: props.allowSelfRemove ? "linear-gradient(135deg,#27AE60,#1E8449)" : "#E2E8F0",
          color: props.allowSelfRemove ? "#fff" : C.muted,
          boxShadow: props.allowSelfRemove ? "0 3px 12px rgba(39,174,96,.3)" : "none",
          minWidth: 120,
        }}>
          {props.allowSelfRemove ? "מופעל" : "מושבת"}
        </button>
      </div>

      {DAYS.map(function(day) {
        var dayShifts = props.shiftMap[day] || [];
        var totalVol = 0;
        for(var i=0;i<dayShifts.length;i++) totalVol += dayShifts[i].maxVol;
        var maxDm = (props.dayConfigs[day]||{maxDayMgr:2}).maxDayMgr;
        var dmCnt = (props.dmOcc[day]||[]).length;

        return (
          <div key={day} style={{background:C.card,borderRadius:16,marginBottom:20,boxShadow:"0 2px 10px rgba(0,0,0,.08)",overflow:"hidden"}}>
            <div style={{background:"#F8FAFC",borderBottom:"2px solid #EEF2F7",padding:"14px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
              <input value={props.dayNames[day]||("יום "+day)}
                onChange={function(e){var v=e.target.value; props.onUpdateDayName(day,v);}}
                style={{padding:"7px 12px",borderRadius:8,border:"2px solid #CBD5E0",fontSize:15,fontWeight:800,color:"#1A202C",background:"#fff",outline:"none",width:240}} />
              <div style={{display:"flex",alignItems:"center",gap:7,background:"#F0FDF9",borderRadius:9,padding:"7px 13px",border:"1px solid #CCFBF1"}}>
                <span style={{fontSize:12,color:C.teal,fontWeight:700}}>📋 אחראי יום:</span>
                <button onClick={function(){if(maxDm>0)props.onUpdateDayConfig(day,maxDm-1);}} disabled={maxDm<=0}
                  style={{width:24,height:24,borderRadius:5,border:"1px solid #CCFBF1",background:"#fff",fontSize:14,cursor:"pointer",fontWeight:700,color:C.teal,display:"flex",alignItems:"center",justifyContent:"center"}}>-</button>
                <span style={{fontSize:17,fontWeight:900,color:C.teal,minWidth:20,textAlign:"center"}}>{maxDm}</span>
                <button onClick={function(){if(maxDm<2)props.onUpdateDayConfig(day,maxDm+1);}} disabled={maxDm>=2}
                  style={{width:24,height:24,borderRadius:5,border:"1px solid #CCFBF1",background:"#fff",fontSize:14,cursor:"pointer",fontWeight:700,color:C.teal,display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
                <span style={{fontSize:10,color:C.muted}}>({dmCnt} נרשמו)</span>
              </div>
              <span style={{fontSize:11,color:C.muted}}>סה"כ: {totalVol} מתנדבים - {dayShifts.length} משמרות</span>
            </div>

            <div style={{padding:"14px 18px"}}>
              {!dayShifts.length && <div style={{textAlign:"center",padding:"16px 0",color:C.muted,fontSize:12,fontStyle:"italic"}}>אין משמרות ליום זה</div>}
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                {dayShifts.map(function(shift) {
                  var regCnt = regCount(shift.id);
                  var canDel = regCnt === 0;
                  var occ = props.occ[shift.id] || {volunteers:[],managers:[]};
                  var pct = shift.maxVol > 0 ? Math.min(1, occ.volunteers.length / shift.maxVol) : 0;
                  var showPicker = iconPicker && iconPicker.day===day && iconPicker.id===shift.id;

                  return (
                    <div key={shift.id} style={{background:"#F8FAFC",borderRadius:12,padding:"12px 16px",border:"1.5px solid #E2E8F0"}}>
                      <div style={{display:"flex",alignItems:"flex-start",gap:10,flexWrap:"wrap"}}>
                        <div style={{position:"relative"}}>
                          <button onClick={function(){setIconPicker(showPicker?null:{day:day,id:shift.id});}}
                            style={{width:44,height:44,borderRadius:9,border:"2px solid #CBD5E0",background:"#fff",fontSize:22,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>
                            {shift.icon}
                          </button>
                          {showPicker && (
                            <div style={{position:"absolute",top:48,right:0,background:"#fff",borderRadius:10,padding:8,boxShadow:"0 8px 28px rgba(0,0,0,.2)",display:"flex",flexWrap:"wrap",gap:3,width:210,zIndex:99}}>
                              {ICONS.map(function(ic){
                                return (
                                  <button key={ic} onClick={function(){props.onUpdateShift(day,shift.id,"icon",ic);setIconPicker(null);}}
                                    style={{width:34,height:34,borderRadius:6,border:ic===shift.icon?"2px solid #2563EB":"1px solid #E2E8F0",background:ic===shift.icon?"#EFF6FF":"#fff",fontSize:18,cursor:"pointer"}}>
                                    {ic}
                                  </button>
                                );
                              })}
                            </div>
                          )}
                        </div>

                        <div style={{flex:1,display:"flex",gap:8,flexWrap:"wrap"}}>
                          <div style={{flex:"1 1 130px"}}>
                            <label style={{fontSize:10,fontWeight:700,color:C.muted,display:"block",marginBottom:3}}>שם המשמרת</label>
                            <input value={shift.name} onChange={function(e){props.onUpdateShift(day,shift.id,"name",e.target.value);}}
                              style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"2px solid #CBD5E0",fontSize:14,fontWeight:700,color:C.navy,outline:"none",boxSizing:"border-box"}} />
                          </div>
                          <div style={{flex:"1 1 130px"}}>
                            <label style={{fontSize:10,fontWeight:700,color:C.muted,display:"block",marginBottom:3}}>שעות / תיאור</label>
                            <input value={shift.hours} onChange={function(e){props.onUpdateShift(day,shift.id,"hours",e.target.value);}}
                              style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"2px solid #CBD5E0",fontSize:12,color:"#1A202C",background:"#fff",outline:"none",boxSizing:"border-box"}} />
                          </div>
                        </div>

                        <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
                          <label style={{fontSize:10,fontWeight:700,color:C.muted}}>מתנדבים</label>
                          <div style={{display:"flex",alignItems:"center",gap:5}}>
                            <button onClick={function(){if(shift.maxVol>0)props.onUpdateShift(day,shift.id,"maxVol",shift.maxVol-1);}}
                              style={{width:26,height:26,borderRadius:6,border:"1px solid #CBD5E0",background:"#fff",fontSize:14,cursor:"pointer",fontWeight:700,color:C.navy}}>-</button>
                            <span style={{fontSize:18,fontWeight:900,color:C.navy,minWidth:32,textAlign:"center"}}>{shift.maxVol}</span>
                            <button onClick={function(){props.onUpdateShift(day,shift.id,"maxVol",shift.maxVol+1);}}
                              style={{width:26,height:26,borderRadius:6,border:"1px solid #CBD5E0",background:"#fff",fontSize:14,cursor:"pointer",fontWeight:700,color:C.navy}}>+</button>
                          </div>
                          <div style={{width:70,height:5,background:"#E2E8F0",borderRadius:3,overflow:"hidden"}}>
                            <div style={{height:"100%",borderRadius:3,width:(pct*100)+"%",background:pct>=1?C.red:pct>=0.7?C.amber:C.green}} />
                          </div>
                          <span style={{fontSize:9,color:C.muted}}>{occ.volunteers.length} נרשמו</span>
                        </div>

                        <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
                          <label style={{fontSize:10,fontWeight:700,color:C.purple}}>אחראים</label>
                          <div style={{display:"flex",alignItems:"center",gap:5}}>
                            <button onClick={function(){if(shift.maxMgr>0)props.onUpdateShift(day,shift.id,"maxMgr",shift.maxMgr-1);}}
                              style={{width:26,height:26,borderRadius:6,border:"1px solid #EDE9FE",background:"#fff",fontSize:14,cursor:"pointer",fontWeight:700,color:C.purple}}>-</button>
                            <span style={{fontSize:18,fontWeight:900,color:C.purple,minWidth:32,textAlign:"center"}}>{shift.maxMgr}</span>
                            <button onClick={function(){props.onUpdateShift(day,shift.id,"maxMgr",shift.maxMgr+1);}}
                              style={{width:26,height:26,borderRadius:6,border:"1px solid #EDE9FE",background:"#fff",fontSize:14,cursor:"pointer",fontWeight:700,color:C.purple}}>+</button>
                          </div>
                          <div style={{width:70,height:5,background:"#E2E8F0",borderRadius:3,overflow:"hidden"}}>
                            <div style={{height:"100%",borderRadius:3,width:Math.min(100,(shift.maxMgr>0?occ.managers.length/shift.maxMgr:0)*100)+"%",background:occ.managers.length>=shift.maxMgr?C.green:C.purple}} />
                          </div>
                          <span style={{fontSize:9,color:C.muted}}>{occ.managers.length} נרשמו</span>
                        </div>

                        <div>
                          {confirmDel===shift.id
                            ? <div style={{background:"#FEF2F2",borderRadius:8,padding:"7px 11px",textAlign:"center"}}>
                                <div style={{fontSize:10,color:C.red,fontWeight:700,marginBottom:5}}>למחוק?</div>
                                <div style={{display:"flex",gap:5}}>
                                  <button onClick={function(){props.onRemoveShift(day,shift.id);setConfirmDel(null);}} style={{padding:"3px 9px",borderRadius:5,border:"none",background:C.red,color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>כן</button>
                                  <button onClick={function(){setConfirmDel(null);}} style={{padding:"3px 9px",borderRadius:5,border:"1px solid #CBD5E0",background:"#fff",fontSize:11,cursor:"pointer",color:C.muted}}>לא</button>
                                </div>
                              </div>
                            : canDel
                              ? <button onClick={function(){setConfirmDel(shift.id);}} style={{background:"#FEF2F2",border:"1px solid "+C.red,color:C.red,borderRadius:7,padding:"7px 11px",fontSize:11,fontWeight:700,cursor:"pointer"}}>מחק</button>
                              : <div style={{background:"#F0F4F8",borderRadius:7,padding:"7px 11px",fontSize:10,color:C.muted,textAlign:"center"}}>{regCnt} נרשמו<br/>לא ניתן</div>
                          }
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              <button onClick={function(){props.onAddShift(day);}}
                style={{marginTop:12,width:"100%",padding:"11px 0",borderRadius:10,border:"2px dashed "+C.navy,background:"transparent",color:C.navy,fontSize:13,fontWeight:700,cursor:"pointer"}}>
                + הוסף משמרת
              </button>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// --- IMPORT ------------------------------------------------------------------
function ImportView(props) {
  var sr = useRef(null);
  var sp = useState(null); var preview = sp[0]; var setPreview = sp[1];
  var ss = useState(null); var status = ss[0]; var setStatus = ss[1];
  var sd = useState(false); var dragging = sd[0]; var setDragging = sd[1];

  function handleFile(file) {
    if (!file) return;
    setPreview(null); setStatus(null);
    parseImport(file,
      function(usersObj, warnings) { setPreview({users:usersObj, warnings:warnings}); },
      function(errors) { setStatus({ok:false, errors:errors}); }
    );
  }

  // Find conflicts: existing user whose type changes AND is registered to a shift/day
  function findConflicts(usersObj) {
    var conflicts = [];
    var ids = Object.keys(usersObj);
    for (var i = 0; i < ids.length; i++) {
      var id = ids[i];
      var newU = usersObj[id];
      var existing = props.users[id];
      if (!existing) continue;
      if (existing.type === newU.type) continue;
      // Type is changing — check if registered
      var shiftId = props.regs[id] || null;
      var dayNum  = props.dmRegs[id] || null;
      if (shiftId || dayNum) {
        var shift = null;
        if (shiftId) {
          for (var j = 0; j < props.shifts.length; j++) {
            if (props.shifts[j].id === shiftId) { shift = props.shifts[j]; break; }
          }
        }
        conflicts.push({
          id: id,
          name: existing.name,
          currentType: existing.type,
          newType: newU.type,
          shiftId: shiftId,
          shiftName: shift ? shift.name : null,
          dayNum: dayNum,
          dayLabel: dayNum ? (props.dayNames[dayNum] || ("יום " + dayNum)) : null,
        });
      }
    }
    return conflicts;
  }

  function commit() {
    if (!preview) return;
    var conflicts = findConflicts(preview.users);
    if (conflicts.length > 0) {
      setStatus({ok:false, conflicts:conflicts});
      setPreview(null);
      return;
    }
    var arr = Object.keys(preview.users).map(function(id) {
      return Object.assign({id:id}, preview.users[id]);
    });
    props.onImport(arr);
    setStatus({ok:true, count:arr.length});
    setPreview(null);
  }

  var previewKeys = preview ? Object.keys(preview.users) : [];

  return (
    <div style={{maxWidth:720}}>
      <div style={{background:C.card,borderRadius:14,padding:"18px 22px",boxShadow:"0 2px 8px rgba(0,0,0,.08)",marginBottom:20}}>
        <h3 style={{color:C.navy,margin:"0 0 5px",fontSize:17,fontWeight:800}}>ייבוא משתמשים מאקסל</h3>
        <p style={{color:C.muted,margin:0,fontSize:12,lineHeight:1.6}}>עמודות: id, type, name, phone, email, password, hr<br/>סוגים: volunteer, manager, day_manager, admin, superadmin<br/>ייבוא כפול: פרטים מתעדכנים, שיבוץ נשמר. שינוי תפקיד — אסור אם המשתמש רשום.</p>
      </div>

      <div style={{border:"2.5px dashed "+(dragging?"#2563EB":"#CBD5E0"),borderRadius:14,padding:"38px 28px",textAlign:"center",cursor:"pointer",background:dragging?"#EFF6FF":"#F8FAFC"}}
        onClick={function(){if(sr.current)sr.current.click();}}
        onDragOver={function(e){e.preventDefault();setDragging(true);}}
        onDragLeave={function(){setDragging(false);}}
        onDrop={function(e){e.preventDefault();setDragging(false);if(e.dataTransfer.files[0])handleFile(e.dataTransfer.files[0]);}}>
        <div style={{fontSize:40,marginBottom:9}}>📂</div>
        <div style={{fontSize:15,fontWeight:700,color:C.navy,marginBottom:4}}>גרור קובץ לכאן, או לחץ לבחירה</div>
        <div style={{fontSize:12,color:C.muted}}>.xlsx .xls .csv</div>
        <input ref={sr} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={function(e){if(e.target.files[0])handleFile(e.target.files[0]);}} />
      </div>

      {preview && (
        <div style={{marginTop:16,background:C.card,borderRadius:14,boxShadow:"0 2px 8px rgba(0,0,0,.08)",overflow:"hidden"}}>
          <div style={{padding:"13px 18px",borderBottom:"1px solid #EEF2F7",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontWeight:800,fontSize:14,color:C.navy}}>{previewKeys.length} משתמשים לייבא</span>
            <button onClick={commit} style={{padding:"7px 18px",borderRadius:9,border:"none",background:"linear-gradient(135deg,#E67E22,#F39C12)",color:"#fff",fontSize:13,fontWeight:800,cursor:"pointer"}}>ייבא</button>
          </div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:460}}>
              <thead><tr style={{background:"#F8FAFC"}}>{["ת.ז.","מצב","סוג חדש","שם","טלפון"].map(function(h,i){return <th key={i} style={{padding:"8px 12px",textAlign:"right",fontWeight:700,color:C.muted,fontSize:11}}>{h}</th>;})}</tr></thead>
              <tbody>
                {previewKeys.slice(0,20).map(function(id,i){
                  var u = preview.users[id];
                  var tm = TYPE_INFO[u.type] || {label:u.type,bg:"#F1F5F9",col:C.muted};
                  var existing = props.users[id];
                  var isNew = !existing;
                  var roleChanged = existing && existing.type !== u.type;
                  var detailsChanged = existing && !roleChanged && (
                    existing.name !== u.name ||
                    (existing.phone||"") !== (u.phone||"") ||
                    (existing.email||"") !== (u.email||"") ||
                    (existing.hr||"") !== (u.hr||"") ||
                    (existing.password||"") !== (u.password||"")
                  );
                  return (
                    <tr key={id} style={{background:i%2===0?"#fff":"#F8FAFC",borderBottom:"1px solid #EEF2F7"}}>
                      <td style={{padding:"8px 12px",fontFamily:"monospace",fontSize:11}}>{id}</td>
                      <td style={{padding:"8px 12px"}}>
                        {isNew
                          ? <span style={{background:"#D1FAE5",color:C.green,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:700}}>חדש</span>
                          : roleChanged
                            ? <span style={{background:"#FEF3C7",color:C.amber,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:700}}>עדכון תפקיד</span>
                            : detailsChanged
                              ? <span style={{background:"#EFF6FF",color:C.blue,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:700}}>עדכון פרטים</span>
                              : <span style={{background:"#F1F5F9",color:C.muted,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:700}}>ללא עדכון</span>
                        }
                      </td>
                      <td style={{padding:"8px 12px"}}><span style={{background:tm.bg,color:tm.col,borderRadius:4,padding:"2px 7px",fontSize:10,fontWeight:700}}>{tm.label}</span></td>
                      <td style={{padding:"8px 12px",fontWeight:600}}>{u.name}</td>
                      <td style={{padding:"8px 12px",color:C.muted}}>{u.phone}</td>
                    </tr>
                  );
                })}
                {previewKeys.length>20 && <tr><td colSpan={5} style={{padding:"8px 12px",textAlign:"center",color:C.muted,fontSize:11}}>+ עוד {previewKeys.length-20}...</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Conflict errors */}
      {status && !status.ok && status.conflicts && (
        <div style={{marginTop:14,background:"#FEF2F2",border:"1.5px solid "+C.red,borderRadius:11,padding:"16px 20px"}}>
          <div style={{fontSize:15,fontWeight:800,color:C.red,marginBottom:12}}>הייבוא נחסם — יש משתמשים רשומים שמנסים לשנות תפקיד:</div>
          {status.conflicts.map(function(c,i){
            var curTm = TYPE_INFO[c.currentType]||{label:c.currentType,bg:"#F1F5F9",col:C.muted};
            var newTm = TYPE_INFO[c.newType]||{label:c.newType,bg:"#F1F5F9",col:C.muted};
            return (
              <div key={i} style={{background:"#fff",borderRadius:9,padding:"12px 16px",marginBottom:8,border:"1px solid #FED7D7"}}>
                <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:6}}>
                  <div style={{fontWeight:800,fontSize:14,color:C.text}}>{c.name}</div>
                  <span style={{fontFamily:"monospace",fontSize:11,color:C.muted}}>ת.ז. {c.id}</span>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:6}}>
                  <span style={{fontSize:12,color:C.muted}}>תפקיד נוכחי:</span>
                  <span style={{background:curTm.bg,color:curTm.col,borderRadius:5,padding:"2px 8px",fontSize:11,fontWeight:700}}>{curTm.label}</span>
                  <span style={{fontSize:12,color:C.muted}}>← תפקיד חדש:</span>
                  <span style={{background:newTm.bg,color:newTm.col,borderRadius:5,padding:"2px 8px",fontSize:11,fontWeight:700}}>{newTm.label}</span>
                </div>
                <div style={{fontSize:12,color:C.red,fontWeight:600}}>
                  {c.shiftName
                    ? "רשום/ה למשמרת: " + c.shiftName + (c.dayLabel ? " (" + c.dayLabel + ")" : "")
                    : "רשום/ה כאחראי/ת יום: " + (c.dayLabel||"")
                  }
                </div>
                <div style={{fontSize:11,color:C.muted,marginTop:4}}>יש להסיר את הרישום לפני שינוי התפקיד.</div>
              </div>
            );
          })}
        </div>
      )}

      {status && !status.ok && status.errors && (
        <div style={{marginTop:14,background:"#FEF2F2",border:"1.5px solid "+C.red,borderRadius:11,padding:"14px 18px"}}>
          <div style={{fontSize:13,fontWeight:800,color:C.red,marginBottom:6}}>שגיאות:</div>
          {status.errors.map(function(e,i){ return <div key={i} style={{fontSize:12,color:C.red}}>- {e}</div>; })}
        </div>
      )}

      {status && status.ok && (
        <div style={{marginTop:14,background:"#D5F5E3",border:"1.5px solid "+C.green,borderRadius:11,padding:"14px 18px"}}>
          <div style={{fontSize:14,fontWeight:800,color:C.green}}>יובאו {status.count} משתמשים בהצלחה!</div>
          <div style={{fontSize:11,color:C.muted,marginTop:3}}>שיבוצים קיימים נשמרו.</div>
        </div>
      )}
    </div>
  );
}

