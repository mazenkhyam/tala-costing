import { useState, useEffect, useRef } from "react";

// ─── storage ──────────────────────────────────────────────────
const sg = (k,fb) => { try { const v=localStorage.getItem(k); return v?JSON.parse(v):fb; } catch { return fb; } };
const ss = (k,v)  => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} };

// ─── helpers ──────────────────────────────────────────────────
const TODAY   = new Date().toISOString().slice(0,10);
const fmt     = d => d ? new Date(d).toLocaleDateString("ar-EG",{day:"2-digit",month:"short",year:"numeric"}) : "—";
const fmtEN   = d => d ? new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}) : "—";
const fmtTime = d => d ? new Date(d).toLocaleString("ar-EG",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"}) : "—";
const diff    = (a,b) => (!a||!b) ? null : Math.round((new Date(b)-new Date(a))/86400000);
const isOD    = t => t.status!=="done" && t.dueDate && new Date(t.dueDate)<new Date();
const uid     = () => Date.now().toString(36)+Math.random().toString(36).slice(2);
const avg     = a => a.length ? Math.round(a.reduce((s,v)=>s+v,0)/a.length) : null;
const tDur    = t => diff(t.startDate,t.dueDate);
const aDur    = t => t.deliveryDate ? diff(t.startDate,t.deliveryDate) : null;
const eff     = t => { const al=tDur(t),ac=aDur(t); return (!al||!ac||al<=0)?null:Math.round(al/ac*100); };
const onTime  = t => t.status==="done"&&t.deliveryDate&&t.dueDate&&new Date(t.deliveryDate)<=new Date(t.dueDate);
const isLate  = t => t.status==="done"&&t.deliveryDate&&t.dueDate&&new Date(t.deliveryDate)>new Date(t.dueDate);

// ─── seed data ────────────────────────────────────────────────
const SEED_USERS = [
  {id:"u1",name:"أحمد محمد", nameEn:"Ahmed Mohammed", email:"admin@tala.com",   password:"admin123",   role:"admin_sys",active:true,avatar:"أم",color:"#c8a97e"},
  {id:"u2",name:"سارة خالد", nameEn:"Sara Khalid",    email:"sara@tala.com",    password:"sara123",    role:"dept_mgr", active:true,avatar:"سخ",color:"#4a9eba"},
  {id:"u3",name:"محمود علي", nameEn:"Mahmoud Ali",    email:"mahmoud@tala.com", password:"mahmoud123", role:"senior",   active:true,avatar:"مع",color:"#5aaa5a"},
  {id:"u4",name:"نورا حسن",  nameEn:"Noura Hassan",   email:"noura@tala.com",   password:"noura123",   role:"junior",   active:true,avatar:"نح",color:"#9a6aba"},
  {id:"u5",name:"ياسر فهد",  nameEn:"Yasser Fahad",   email:"yasser@tala.com",  password:"yasser123",  role:"junior",   active:true,avatar:"يف",color:"#c8604a"},
];
const SEED_TASKS = [];
const SEED_PERMS = {
  admin_sys:{canAdd:true, canAssign:true, canViewAll:true, canManageUsers:true, canReports:true, canApprove:true, canEditRoles:true},
  dept_mgr: {canAdd:true, canAssign:true, canViewAll:true, canManageUsers:false,canReports:true, canApprove:true, canEditRoles:false},
  senior:   {canAdd:true, canAssign:false,canViewAll:false,canManageUsers:false,canReports:false,canApprove:false,canEditRoles:false},
  junior:   {canAdd:false,canAssign:false,canViewAll:false,canManageUsers:false,canReports:false,canApprove:false,canEditRoles:false},
};
const ROLE_LABEL = {
  ar:{admin_sys:"مدير النظام",dept_mgr:"مدير قسم",senior:"محاسب أول",junior:"محاسب"},
  en:{admin_sys:"System Admin",dept_mgr:"Dept. Manager",senior:"Sr. Accountant",junior:"Accountant"},
};
const ROLE_COLOR = {admin_sys:"#c8a97e",dept_mgr:"#4a9eba",senior:"#5aaa5a",junior:"#9a6aba"};
const S_CFG = {
  pending:   {ar:"انتظار",      en:"Pending",     icon:"◷",color:"#8888aa",dbg:"rgba(136,136,170,.12)",dlg:"rgba(80,80,120,.12)"},
  inprogress:{ar:"تحت الإجراء",en:"In Progress",  icon:"⟳",color:"#4a9eba",dbg:"rgba(74,158,186,.12)", dlg:"rgba(30,100,140,.12)"},
  blocked:   {ar:"عوائق",       en:"Blocked",      icon:"⚠",color:"#e07060",dbg:"rgba(224,112,96,.12)", dlg:"rgba(180,60,50,.12)"},
  done:      {ar:"مكتملة",      en:"Completed",    icon:"✓",color:"#5aaa5a",dbg:"rgba(90,170,90,.12)",  dlg:"rgba(40,120,50,.12)"},
};
const PC = {urgent:"#e07060",high:"#c8a97e",medium:"#4a9eba",low:"#888898"};

// ─── CSS VARIABLES theme ──────────────────────────────────────
// All UI colors come from CSS vars — toggling .tala-dark / .tala-light swaps everything
const THEME_CSS = `
  .tala-dark {
    --bg0:#07070d; --bg1:#0c0c16; --bg2:#10101e; --bg3:#161626; --bg4:#1c1c30;
    --border:#1a1a2e; --border2:#26263c;
    --text:#ffffff; --mid:#d0d0e8; --dim:#9090b8;
    --card-bg:#10101e; --card-border:#1a1a2e;
    --input-bg:#07070d; --input-border:#26263c; --input-color:#ffffff;
    --sidebar-bg:#0c0c16; --topbar-bg:#0c0c16;
    --modal-bg:#0c0c16;
    --hover-bg:#1c1c30;
    --table-alt:#0e0e1c;
    --shadow:0 24px 80px rgba(0,0,0,.6);
    --note-bg:#161626;
    --badge-bg:#1c1c30;
    --gold:#c8a97e; --teal:#5cb8b2;
  }
  .tala-light {
    --bg0:#edeae4; --bg1:#ffffff; --bg2:#f7f5f1; --bg3:#ece8e1; --bg4:#e2ddd6;
    --border:#d4cfc8; --border2:#b8b2aa;
    --text:#000000; --mid:#111111; --dim:#333333;
    --card-bg:#ffffff; --card-border:#d4cfc8;
    --input-bg:#f0ece6; --input-border:#b8b2aa; --input-color:#000000;
    --sidebar-bg:#f7f5f1; --topbar-bg:#ffffff;
    --modal-bg:#ffffff;
    --hover-bg:#e8e3dc;
    --table-alt:#f2efe9;
    --shadow:0 16px 60px rgba(0,0,0,.14);
    --note-bg:#ece8e1;
    --badge-bg:#e8e3dc;
    --gold:#7a5020; --teal:#1a7870;
  }

  * { box-sizing:border-box; margin:0; padding:0; }
  *:not(svg):not(path):not(circle) { color:inherit; }
  body, input, select, textarea, button { font-family:inherit; color:inherit; }
  ::-webkit-scrollbar { width:4px; height:4px; }
  ::-webkit-scrollbar-track { background:var(--bg0); }
  ::-webkit-scrollbar-thumb { background:var(--border2); border-radius:3px; }
  select option { background:var(--bg2); color:var(--text); }
  input[type=date]::-webkit-calendar-picker-indicator { opacity:.5; cursor:pointer; }
  button:focus { outline:none; }
  .tala-dark  input[type=date]::-webkit-calendar-picker-indicator { filter:invert(.6); }
  .tala-light input[type=date]::-webkit-calendar-picker-indicator { filter:invert(.2); }
  /* Headings always use full text color */
  h1,h2,h3,h4 { color:var(--text) !important; font-family:inherit; }
  /* All p/span inherit unless overridden */
  p, span, td, th, label { font-family:inherit; }
  @keyframes fadeUp { from{opacity:0;transform:translateY(16px)} to{opacity:1;transform:translateY(0)} }
  @keyframes shake  { 0%,100%{transform:translateX(0)} 20%,60%{transform:translateX(-7px)} 40%,80%{transform:translateX(7px)} }
  @keyframes spin   { to{transform:rotate(360deg)} }
`;

// ─── dynamic style helpers (use CSS vars) ─────────────────────
const V = k => `var(--${k})`;
const inp  = { background:V("input-bg"), border:`1px solid ${V("input-border")}`, color:V("input-color"), borderRadius:8, padding:"9px 13px", fontSize:13, outline:"none", fontFamily:"inherit", width:"100%", boxSizing:"border-box", transition:"border-color .2s" };
const sel  = { ...inp, cursor:"pointer" };
const lbl  = { color:V("mid"), fontSize:10, fontWeight:700, letterSpacing:".08em", textTransform:"uppercase", display:"block", marginBottom:5, fontFamily:"inherit" };
const CARD = { background:V("card-bg"), border:`1px solid ${V("card-border")}`, borderRadius:10 };
const mkBtnG = () => ({ background:`linear-gradient(135deg,var(--gold),#a87a50)`, color:"#0a0806", border:"none", borderRadius:8, padding:"9px 20px", cursor:"pointer", fontWeight:700, fontFamily:"inherit", fontSize:13 });
const mkBtnO = () => ({ background:"none", color:V("mid"), border:`1px solid ${V("border2")}`, borderRadius:8, padding:"7px 15px", cursor:"pointer", fontFamily:"inherit", fontSize:12 });

// ─── i18n ─────────────────────────────────────────────────────
const TR = {
  ar:{appName:"تالا",appSub:"تتبع الأعمال واللوائح المحاسبية",dashboard:"لوحة الأداء",tasks:"المهام",users:"المستخدمون",perms:"الصلاحيات",myTasks:"مهامي",allTasks:"جميع المهام",addTask:"+ مهمة جديدة",save:"حفظ",cancel:"إلغاء",del:"حذف",logout:"تسجيل الخروج",login:"تسجيل الدخول",email:"البريد الإلكتروني",pass:"كلمة المرور",name:"الاسم",role:"الدور",status:"الحالة",assignTo:"مسند إلى",priority:"الأولوية",startD:"تاريخ البداية",dueD:"الموعد النهائي",delivD:"تاريخ التسليم",registered:"تاريخ التسجيل",allocDays:"أيام الإنجاز",notes:"الملاحظات",addNote:"إضافة",blocker:"وصف العائق",completeBtn:"✓ إنجاز المهمة",confirmTitle:"تأكيد إنجاز المهمة",reqCorr:"طلب تصحيح التاريخ",corrReason:"سبب التصحيح",sendReq:"إرسال",approveBtn:"✓ موافقة",rejectBtn:"✕ رفض",corrPanel:"طلبات تصحيح تاريخ التسليم",total:"الإجمالي",done:"مكتملة",prog:"تحت الإجراء",blocked:"عوائق",overdue:"متأخرة",pending:"انتظار",onTimeL:"في الموعد",lateL:"متأخرة",allocL:"المدة المخصصة",actualL:"مدة الإنجاز",effL:"كفاءة",compRate:"نسبة الإنجاز",otRate:"الالتزام بالمواعيد",avgAlloc:"متوسط الأيام المخصصة",avgActual:"متوسط أيام الإنجاز",avgEff:"متوسط الكفاءة",empPerf:"أداء الموظفين",aiBtn:"تحليل AI",aiTitle:"تحليل الذكاء الاصطناعي",analyzing:"جاري التحليل...",exportBtn:"⬇ تصدير Excel",dayL:"يوم",autoCalc:"يُحسب تلقائياً",taskTitle:"عنوان المهمة",allEmp:"جميع الموظفين",allStatus:"الكل",boardV:"بورد",listV:"قائمة",noTasks:"لا توجد مهام",saved:"محفوظ",notifs:"الإشعارات",newN:"جديد",markAll:"قراءة الكل",noNotifs:"لا توجد إشعارات",addUser:"+ إضافة مستخدم",newUser:"مستخدم جديد",fullname:"الاسم الكامل",editU:"تعديل",disableU:"تعطيل",enableU:"تفعيل",permTitle:"إدارة الصلاحيات",permDesc:"تحكم في ما يستطيع كل دور القيام به",canAdd:"إضافة مهام",canAssign:"إسناد للآخرين",canViewAll:"عرض كل المهام",canManageUsers:"إدارة المستخدمين",canReports:"عرض التقارير",canApprove:"الموافقة على التصحيحات",canEditRoles:"تعديل الصلاحيات",colorL:"لون التمييز",urgent:"عاجلة",high:"عالية",medium:"متوسطة",low:"منخفضة",demoHint:"بيانات تجريبية — اضغط للملء",deliverLock:"سيُسجَّل تاريخ التسليم ولا يمكن تعديله إلا بموافقة المدير",noFuture:"الافتراضي اليوم — لا يمكن اختيار تاريخ مستقبلي",corrWarn:"سيُرسل الطلب للمدير للموافقة",currentDateL:"التاريخ الحالي",correctDateL:"التاريخ الصحيح",reqBy:"بواسطة",reqDate:"التاريخ المطلوب",reasonL:"السبب",detailsL:"تفاصيل المهام المنجزة",taskOf:"المهمة",darkMode:"الوضع الداكن",lightMode:"الوضع الفاتح"},
  en:{appName:"TALA",appSub:"Task & Accounting Ledger Analytics",dashboard:"Dashboard",tasks:"Tasks",users:"Users",perms:"Permissions",myTasks:"My Tasks",allTasks:"All Tasks",addTask:"+ New Task",save:"Save",cancel:"Cancel",del:"Delete",logout:"Sign Out",login:"Sign In",email:"Email",pass:"Password",name:"Name",role:"Role",status:"Status",assignTo:"Assigned To",priority:"Priority",startD:"Start Date",dueD:"Due Date",delivD:"Delivery Date",registered:"Registered",allocDays:"Allocated Days",notes:"Notes",addNote:"Add",blocker:"Blocker",completeBtn:"✓ Mark Complete",confirmTitle:"Confirm Task Completion",reqCorr:"Request Date Correction",corrReason:"Reason",sendReq:"Send",approveBtn:"✓ Approve",rejectBtn:"✕ Reject",corrPanel:"Delivery Date Correction Requests",total:"Total",done:"Completed",prog:"In Progress",blocked:"Blocked",overdue:"Overdue",pending:"Pending",onTimeL:"On Time",lateL:"Late",allocL:"Allocated",actualL:"Actual",effL:"Efficiency",compRate:"Completion",otRate:"On-Time Rate",avgAlloc:"Avg Alloc Days",avgActual:"Avg Actual Days",avgEff:"Avg Efficiency",empPerf:"Employee Performance",aiBtn:"AI Analysis",aiTitle:"AI Analysis",analyzing:"Analyzing...",exportBtn:"⬇ Export Excel",dayL:"day",autoCalc:"Auto-calculated",taskTitle:"Task Title",allEmp:"All Employees",allStatus:"All",boardV:"Board",listV:"List",noTasks:"No tasks found",saved:"Saved",notifs:"Notifications",newN:"new",markAll:"Mark all read",noNotifs:"No notifications",addUser:"+ Add User",newUser:"New User",fullname:"Full Name",editU:"Edit",disableU:"Disable",enableU:"Enable",permTitle:"Permissions",permDesc:"Control what each role can do",canAdd:"Add Tasks",canAssign:"Assign to Others",canViewAll:"View All Tasks",canManageUsers:"Manage Users",canReports:"View Reports",canApprove:"Approve Corrections",canEditRoles:"Edit Roles",colorL:"Color",urgent:"Urgent",high:"High",medium:"Medium",low:"Low",demoHint:"Demo credentials — click to fill",deliverLock:"Delivery date is locked. Changes require admin approval.",noFuture:"Default is today — future dates not allowed",corrWarn:"Request will be sent to manager for approval",currentDateL:"Current Date",correctDateL:"Correct Date",reqBy:"By",reqDate:"Requested Date",reasonL:"Reason",detailsL:"Completed Task Details",taskOf:"Task",darkMode:"Dark Mode",lightMode:"Light Mode"},
};

// ═══════════════════════════════════════════════════════════════
//  THEME TOGGLE BUTTON
// ═══════════════════════════════════════════════════════════════
function ThemeBtn({ dark, onToggle, t }) {
  return (
    <button onClick={onToggle} title={dark ? t.lightMode : t.darkMode}
      style={{ background:"none", border:`1px solid ${V("border2")}`, borderRadius:8, width:36, height:36, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", fontSize:16, color:V("mid"), transition:"all .2s" }}
      onMouseEnter={e=>e.currentTarget.style.borderColor=V("gold")}
      onMouseLeave={e=>e.currentTarget.style.borderColor=V("border2")}>
      {dark ? "☀️" : "🌙"}
    </button>
  );
}

// ═══════════════════════════════════════════════════════════════
//  CIRCLE STAT
// ═══════════════════════════════════════════════════════════════
function CircleStat({ pct, color, size=64, label, sub }) {
  const stroke=5, r=(size-stroke*2)/2, circ=2*Math.PI*r;
  const offset = circ-(Math.min(100,Math.max(0,pct||0))/100)*circ;
  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
      <div style={{position:"relative",width:size,height:size}}>
        <svg width={size} height={size} style={{transform:"rotate(-90deg)",position:"absolute"}}>
          <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={V("border2")} strokeWidth={stroke}/>
          <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={color} strokeWidth={stroke}
            strokeDasharray={circ} strokeDashoffset={offset} strokeLinecap="round"
            style={{transition:"stroke-dashoffset .8s ease"}}/>
        </svg>
        <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <span style={{color,fontSize:size>60?13:11,fontWeight:800}}>{label}</span>
        </div>
      </div>
      {sub&&<span style={{color:V("dim"),fontSize:9,textAlign:"center",maxWidth:70,lineHeight:1.3}}>{sub}</span>}
    </div>
  );
}

function Bar({pct,color,h=4}){
  return <div style={{background:V("border"),borderRadius:3,height:h,overflow:"hidden"}}><div style={{height:"100%",width:`${Math.min(100,Math.max(0,pct||0))}%`,background:color,borderRadius:3,transition:"width .7s ease"}}/></div>;
}

// ═══════════════════════════════════════════════════════════════
//  LOGIN
// ═══════════════════════════════════════════════════════════════
function Login({ users, onLogin, lang, setLang, dark, setDark }) {
  const [email,setEmail]=useState("");
  const [pass,setPass]=useState("");
  const [err,setErr]=useState("");
  const [shake,setShake]=useState(false);
  const t=TR[lang];

  const submit = () => {
    const u=users.find(x=>x.email===email.trim()&&x.password===pass&&x.active);
    if(u) onLogin(u);
    else { setErr(lang==="ar"?"البريد أو كلمة المرور غير صحيحة":"Invalid email or password"); setShake(true); setTimeout(()=>setShake(false),500); }
  };

  return (
    <div style={{minHeight:"100vh",background:V("bg0"),display:"flex",alignItems:"center",justifyContent:"center",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",inset:0,backgroundImage:`linear-gradient(${V("border")} 1px,transparent 1px),linear-gradient(90deg,${V("border")} 1px,transparent 1px)`,backgroundSize:"48px 48px",opacity:.5}}/>
      <div style={{position:"absolute",top:"15%",right:"8%",width:360,height:360,background:`radial-gradient(circle,var(--gold)18 0%,transparent 65%)`}}/>
      <div style={{position:"absolute",bottom:"15%",left:"8%",width:240,height:240,background:`radial-gradient(circle,var(--teal)14 0%,transparent 65%)`}}/>

      {/* top controls */}
      <div style={{position:"absolute",top:16,right:16,display:"flex",gap:8,zIndex:10}}>
        <button onClick={()=>setLang(l=>l==="ar"?"en":"ar")} style={{...mkBtnO(),fontSize:11,padding:"3px 10px",color:V("gold"),borderColor:`var(--gold)44`}}>{lang==="ar"?"EN":"ع"}</button>
        <ThemeBtn dark={dark} onToggle={()=>setDark(d=>!d)} t={t}/>
      </div>

      <div style={{...CARD,width:420,padding:44,position:"relative",zIndex:1,animation:shake?"shake .4s":"fadeUp .5s",boxShadow:V("shadow")}}>
        {/* logo */}
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{width:68,height:68,margin:"0 auto 16px",borderRadius:18,background:dark?"linear-gradient(135deg,#1e1810,#161214)":"linear-gradient(135deg,#faf5ec,#f0e8d8)",border:`1px solid var(--gold)55`,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <svg viewBox="0 0 40 40" fill="none" style={{width:40,height:40}}>
              <path d="M8 32 L20 6 L32 32" stroke="var(--gold)" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
              <path d="M12 24 L28 24" stroke="var(--gold)" strokeWidth="2" strokeLinecap="round" opacity=".5"/>
              <circle cx="20" cy="34.5" r="2" fill="var(--teal)"/>
            </svg>
          </div>
          <h1 style={{color:V("gold"),margin:"0 0 5px",fontSize:30,fontWeight:800,letterSpacing:"2px"}}>{t.appName}</h1>
          <p style={{color:V("dim"),margin:0,fontSize:11,letterSpacing:".04em"}}>{t.appSub}</p>
        </div>

        <div style={{marginBottom:13}}><label style={lbl}>{t.email}</label><input value={email} onChange={e=>{setEmail(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&submit()} style={inp} placeholder="email@tala.com" dir="ltr"/></div>
        <div style={{marginBottom:10}}><label style={lbl}>{t.pass}</label><input type="password" value={pass} onChange={e=>{setPass(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&submit()} style={inp} placeholder="••••••••" dir="ltr"/></div>
        {err&&<p style={{color:"#e07060",fontSize:12,margin:"6px 0 0",textAlign:"center"}}>{err}</p>}
        <button onClick={submit} style={{...mkBtnG(),width:"100%",marginTop:20,fontSize:14,padding:13}}>{t.login}</button>


      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  NOTIFICATION BELL
// ═══════════════════════════════════════════════════════════════
function NotifBell({ notifs, onRead, onReadAll, t }) {
  const [open,setOpen]=useState(false);
  const ref=useRef(null);
  const unread=notifs.filter(n=>!n.read).length;
  const ICONS={task_assigned:"📋",task_completed:"✅",correction_requested:"🔔",correction_approved:"✓",correction_rejected:"✕",task_blocked:"⚠"};
  useEffect(()=>{ const fn=e=>{ if(ref.current&&!ref.current.contains(e.target)) setOpen(false); }; document.addEventListener("mousedown",fn); return ()=>document.removeEventListener("mousedown",fn); },[]);
  return (
    <div ref={ref} style={{position:"relative"}}>
      <button onClick={()=>setOpen(o=>!o)} style={{position:"relative",background:V("bg3"),border:`1px solid ${V("border")}`,borderRadius:9,width:36,height:36,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,transition:"all .2s"}}>
        🔔
        {unread>0&&<span style={{position:"absolute",top:-5,right:-5,background:"#e07060",color:"white",borderRadius:"50%",width:17,height:17,fontSize:9,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,border:`2px solid ${V("bg0")}`}}>{unread>9?"9+":unread}</span>}
      </button>
      {open&&(
        <div style={{position:"absolute",top:44,left:0,width:310,background:V("modal-bg"),border:`1px solid ${V("border2")}`,borderRadius:12,boxShadow:V("shadow"),zIndex:400,overflow:"hidden"}}>
          <div style={{padding:"10px 14px",borderBottom:`1px solid ${V("border")}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{color:V("text"),fontWeight:700,fontSize:13}}>{t.notifs} {unread>0&&<span style={{color:"#e07060",fontSize:11}}>({unread} {t.newN})</span>}</span>
            {unread>0&&<button onClick={onReadAll} style={{...mkBtnO(),padding:"2px 8px",fontSize:10,color:V("gold"),borderColor:`var(--gold)44`}}>{t.markAll}</button>}
          </div>
          <div style={{maxHeight:320,overflowY:"auto"}}>
            {notifs.length===0&&<p style={{color:V("dim"),textAlign:"center",padding:"22px 0",fontSize:12}}>{t.noNotifs}</p>}
            {notifs.map(n=>(
              <div key={n.id} onClick={()=>onRead(n.id)}
                style={{padding:"10px 14px",borderBottom:`1px solid ${V("border")}`,background:n.read?"transparent":V("bg3"),cursor:"pointer",display:"flex",gap:8,alignItems:"flex-start"}}
                onMouseEnter={e=>e.currentTarget.style.background=V("hover-bg")}
                onMouseLeave={e=>e.currentTarget.style.background=n.read?"transparent":V("bg3")}>
                <span style={{fontSize:14,flexShrink:0}}>{ICONS[n.type]||"📌"}</span>
                <div style={{flex:1}}>
                  <p style={{color:n.read?V("mid"):V("text"),fontSize:12,margin:"0 0 2px",lineHeight:1.5,fontWeight:n.read?400:600}}>{n.text}</p>
                  <p style={{color:V("dim"),fontSize:10,margin:0}}>{fmtTime(n.createdAt)}</p>
                </div>
                {!n.read&&<div style={{width:6,height:6,borderRadius:"50%",background:"#e07060",flexShrink:0,marginTop:4}}/>}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  MODAL WRAPPER
// ═══════════════════════════════════════════════════════════════
function Modal({ onClose, children, width="min(660px,96vw)", borderColor="" }) {
  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:500,backdropFilter:"blur(6px)"}}>
      <div style={{background:V("modal-bg"),border:`1px solid ${borderColor||V("border2")}`,borderRadius:14,width,maxHeight:"92vh",overflowY:"auto",padding:28,position:"relative",boxShadow:V("shadow"),animation:"fadeUp .3s ease"}}>
        <button onClick={onClose} style={{position:"absolute",top:13,left:13,background:V("bg3"),border:`1px solid ${V("border")}`,color:V("mid"),borderRadius:7,width:29,height:29,cursor:"pointer",fontSize:12}}>✕</button>
        {children}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  COMPLETE MODAL
// ═══════════════════════════════════════════════════════════════
function CompleteModal({ task, onClose, onConfirm, t, lang }) {
  const [d,setD]=useState(TODAY);
  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:700,backdropFilter:"blur(8px)"}}>
      <div style={{background:V("modal-bg"),border:"1px solid rgba(90,170,90,.4)",borderRadius:14,width:"min(400px,96vw)",padding:30,boxShadow:V("shadow"),animation:"fadeUp .3s ease"}}>
        <div style={{textAlign:"center",marginBottom:20}}>
          <div style={{fontSize:38,marginBottom:10}}>✅</div>
          <h2 style={{color:V("text"),margin:"0 0 6px",fontSize:17,fontWeight:700}}>{t.confirmTitle}</h2>
          <p style={{color:V("mid"),margin:0,fontSize:12,lineHeight:1.7}}>{t.deliverLock}</p>
        </div>
        <div style={{background:V("bg3"),border:`1px solid ${V("border")}`,borderRadius:9,padding:14,marginBottom:16}}>
          <p style={{color:V("mid"),margin:"0 0 3px",fontSize:10}}>{t.taskOf}</p>
          <p style={{color:V("text"),margin:0,fontSize:13,fontWeight:600}}>{lang==="ar"?task.title:task.titleEn||task.title}</p>
        </div>
        <div style={{marginBottom:18}}>
          <label style={lbl}>{t.delivD}</label>
          <input type="date" value={d} onChange={e=>setD(e.target.value)} max={TODAY} style={inp}/>
          <p style={{color:V("dim"),fontSize:10,margin:"5px 0 0"}}>{t.noFuture}</p>
        </div>
        <div style={{display:"flex",gap:9,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={mkBtnO()}>{t.cancel}</button>
          <button onClick={()=>onConfirm(d)} style={{...mkBtnG(),background:"#5aaa5a",color:"#071207"}}>{t.completeBtn}</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  CORRECTION MODAL
// ═══════════════════════════════════════════════════════════════
function CorrModal({ task, onClose, onReq, t }) {
  const [d,setD]=useState(task.deliveryDate||TODAY);
  const [reason,setReason]=useState("");
  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:700,backdropFilter:"blur(8px)"}}>
      <div style={{background:V("modal-bg"),border:"1px solid rgba(200,169,126,.4)",borderRadius:14,width:"min(420px,96vw)",padding:30,boxShadow:V("shadow"),animation:"fadeUp .3s ease"}}>
        <h2 style={{color:V("text"),margin:"0 0 5px",fontSize:16,fontWeight:700}}>{t.reqCorr}</h2>
        <p style={{color:V("mid"),margin:"0 0 20px",fontSize:12}}>{t.corrWarn}</p>
        <div style={{background:V("bg3"),border:`1px solid ${V("border")}`,borderRadius:9,padding:"10px 13px",marginBottom:13}}>
          <p style={{color:V("mid"),margin:"0 0 2px",fontSize:10}}>{t.currentDateL}</p>
          <p style={{color:"#5aaa5a",margin:0,fontSize:13,fontWeight:600}}>{fmt(task.deliveryDate)}</p>
        </div>
        <div style={{marginBottom:13}}><label style={lbl}>{t.correctDateL}</label><input type="date" value={d} onChange={e=>setD(e.target.value)} max={TODAY} style={inp}/></div>
        <div style={{marginBottom:18}}><label style={lbl}>{t.corrReason}</label><textarea rows={3} value={reason} onChange={e=>setReason(e.target.value)} style={{...inp,resize:"vertical"}} placeholder={t.corrReason+"..."}/></div>
        <div style={{display:"flex",gap:9,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={mkBtnO()}>{t.cancel}</button>
          <button onClick={()=>{ if(!reason.trim())return; onReq({requestedDate:d,reason,taskId:task.id,taskTitle:task.title,requestedBy:task.assignedTo,status:"pending",id:uid(),createdAt:new Date().toISOString()}); onClose(); }} style={mkBtnG()}>{t.sendReq}</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  TASK MODAL
// ═══════════════════════════════════════════════════════════════
function TaskModal({ task, users, me, perms, onClose, onSave, onDelete, onComplete, onCorrReq, t, lang }) {
  const [f,setF]=useState({...task,notes:[...task.notes]});
  const [note,setNote]=useState("");
  const [showComp,setShowComp]=useState(false);
  const [showCorr,setShowCorr]=useState(false);
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  const isOwner=me.id===task.assignedTo;
  const canEdit=perms.canViewAll||isOwner;
  const addNote=()=>{ if(!note.trim())return; set("notes",[...f.notes,{author:me.id,text:note.trim(),date:TODAY}]); setNote(""); };
  const s=S_CFG[f.status];
  const al=tDur(f); const ac=aDur(f); const ev=eff(f);
  const pL={urgent:t.urgent,high:t.high,medium:t.medium,low:t.low};

  return (
    <>
    <Modal onClose={onClose} width="min(700px,96vw)">
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16,paddingLeft:36}}>
        <div>
          <h2 style={{color:V("text"),margin:"0 0 4px",fontSize:15,fontWeight:700,lineHeight:1.5}}>{lang==="ar"?f.title:f.titleEn||f.title}</h2>
          <p style={{color:V("dim"),margin:0,fontSize:10}}>{t.registered}: {lang==="ar"?fmt(f.createdAt):fmtEN(f.createdAt)}</p>
        </div>
        <span style={{color:s.color,background:dark_?s.dbg:s.dlg,border:`1px solid ${s.color}55`,borderRadius:20,padding:"2px 10px",fontSize:11,whiteSpace:"nowrap",marginRight:8}}>{s.icon} {lang==="ar"?s.ar:s.en}</span>
      </div>

      {(al||ac!=null)&&(
        <div style={{display:"flex",gap:9,marginBottom:16,flexWrap:"wrap"}}>
          {al&&<div style={{background:V("bg3"),border:`1px solid ${V("border")}`,borderRadius:9,padding:"8px 13px",flex:1,minWidth:90}}><p style={{color:V("dim"),fontSize:9,margin:"0 0 2px"}}>{t.allocL}</p><p style={{color:V("gold"),fontSize:18,fontWeight:800,margin:0}}>{al}<span style={{fontSize:10,fontWeight:400}}> {t.dayL}</span></p></div>}
          {ac!=null&&<div style={{background:V("bg3"),border:`1px solid ${V("border")}`,borderRadius:9,padding:"8px 13px",flex:1,minWidth:90}}><p style={{color:V("dim"),fontSize:9,margin:"0 0 2px"}}>{t.actualL}</p><p style={{color:ac<=al?"#5aaa5a":"#e07060",fontSize:18,fontWeight:800,margin:0}}>{ac}<span style={{fontSize:10,fontWeight:400}}> {t.dayL}</span></p></div>}
          {ev!=null&&<div style={{background:V("bg3"),border:`1px solid ${V("border")}`,borderRadius:9,padding:"8px 13px",flex:1,minWidth:90}}><p style={{color:V("dim"),fontSize:9,margin:"0 0 2px"}}>{t.effL}</p><p style={{color:ev>=100?"#5aaa5a":"#e07060",fontSize:18,fontWeight:800,margin:0}}>{ev}%</p></div>}
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11,marginBottom:14}}>
        <div><label style={lbl}>{t.status}</label><select value={f.status} onChange={e=>set("status",e.target.value)} disabled={!canEdit||f.status==="done"} style={sel}>{Object.entries(S_CFG).map(([k,v])=><option key={k} value={k}>{v.icon} {lang==="ar"?v.ar:v.en}</option>)}</select></div>
        <div><label style={lbl}>{t.assignTo}</label><select value={f.assignedTo} onChange={e=>set("assignedTo",e.target.value)} disabled={!perms.canAssign} style={sel}>{users.filter(u=>u.active).map(u=><option key={u.id} value={u.id}>{lang==="ar"?u.name:u.nameEn}</option>)}</select></div>
        <div><label style={lbl}>{t.priority}</label><select value={f.priority} onChange={e=>set("priority",e.target.value)} disabled={!canEdit} style={sel}>{Object.entries(pL).map(([k,v])=><option key={k} value={k}>{v}</option>)}</select></div>
        <div><label style={lbl}>{t.allocDays}</label><input type="number" min={1} max={365} value={f.allocDays||""} onChange={e=>set("allocDays",+e.target.value)} disabled={!perms.canAssign} style={inp}/></div>
        <div><label style={lbl}>{t.startD}</label><input type="date" value={f.startDate||""} onChange={e=>set("startDate",e.target.value)} disabled={!perms.canAssign} style={inp}/></div>
        <div><label style={lbl}>{t.dueD}</label><input type="date" value={f.dueDate||""} onChange={e=>set("dueDate",e.target.value)} disabled={!perms.canAssign} style={inp}/></div>
        <div><label style={lbl}>{t.delivD}</label>
          {perms.canViewAll
            ?<input type="date" value={f.deliveryDate||""} onChange={e=>set("deliveryDate",e.target.value)} style={inp}/>
            :<div style={{...inp,color:f.deliveryDate?"#5aaa5a":V("dim"),cursor:"default"}}>{f.deliveryDate?(lang==="ar"?fmt(f.deliveryDate):fmtEN(f.deliveryDate)):(lang==="ar"?"لم يُسلَّم بعد":"Not delivered yet")}</div>
          }
        </div>
      </div>

      {(f.status==="blocked"||f.blockerNote)&&<div style={{marginBottom:13}}><label style={lbl}>{t.blocker}</label><textarea rows={2} value={f.blockerNote} onChange={e=>set("blockerNote",e.target.value)} disabled={!canEdit} style={{...inp,resize:"vertical"}} placeholder={t.blocker+"..."}/></div>}

      <div>
        <label style={{...lbl,marginBottom:8}}>{t.notes} ({f.notes.length})</label>
        <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:10,maxHeight:140,overflowY:"auto"}}>
          {f.notes.length===0&&<p style={{color:V("dim"),fontSize:12,margin:0}}>{lang==="ar"?"لا توجد ملاحظات":"No notes yet"}</p>}
          {f.notes.map((n,i)=>{ const u=users.find(x=>x.id===n.author); return (
            <div key={i} style={{background:V("note-bg"),border:`1px solid ${V("border")}`,borderRadius:8,padding:"8px 12px"}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                <span style={{color:u?.color,fontSize:11,fontWeight:600}}>{lang==="ar"?u?.name:u?.nameEn||u?.name}</span>
                <span style={{color:V("dim"),fontSize:10}}>{lang==="ar"?fmt(n.date):fmtEN(n.date)}</span>
              </div>
              <p style={{color:V("text"),margin:0,fontSize:12,lineHeight:1.6,opacity:.85}}>{n.text}</p>
            </div>);
          })}
        </div>
        <div style={{display:"flex",gap:7}}>
          <input value={note} onChange={e=>setNote(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addNote()} style={{...inp,flex:1}} placeholder={lang==="ar"?"أضف ملاحظة...":"Add a note..."}/>
          <button onClick={addNote} style={{...mkBtnG(),padding:"8px 15px"}}>{t.addNote}</button>
        </div>
      </div>

      <div style={{display:"flex",justifyContent:"space-between",marginTop:18,flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",gap:7}}>
          {perms.canViewAll&&<button onClick={()=>{if(window.confirm(lang==="ar"?"حذف المهمة؟":"Delete task?")){onDelete(task.id);onClose();}}} style={{...mkBtnO(),color:"#e07060",borderColor:"rgba(224,112,96,.3)"}}>{t.del}</button>}
          {isOwner&&!perms.canViewAll&&f.status!=="done"&&<button onClick={()=>setShowComp(true)} style={{...mkBtnG(),background:"#5aaa5a",color:"#071207",fontSize:12}}>{t.completeBtn}</button>}
          {isOwner&&!perms.canViewAll&&f.status==="done"&&f.deliveryDate&&<button onClick={()=>setShowCorr(true)} style={{...mkBtnO(),fontSize:11,color:V("gold"),borderColor:`var(--gold)44`}}>{t.reqCorr}</button>}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={onClose} style={mkBtnO()}>{t.cancel}</button>
          {canEdit&&<button onClick={()=>{onSave(f);onClose();}} style={mkBtnG()}>{t.save}</button>}
        </div>
      </div>
    </Modal>
    {showComp&&<CompleteModal task={task} onClose={()=>setShowComp(false)} onConfirm={d=>{onComplete(task.id,d);onClose();}} t={t} lang={lang}/>}
    {showCorr&&<CorrModal task={task} onClose={()=>setShowCorr(false)} onReq={req=>{onCorrReq(req);onClose();}} t={t}/>}
    </>
  );
}
// marker for dark access
let dark_ = true;

// ═══════════════════════════════════════════════════════════════
//  ADD TASK MODAL
// ═══════════════════════════════════════════════════════════════
function AddTaskModal({ users, me, onClose, onAdd, t, lang }) {
  const [f,setF]=useState({title:"",titleEn:"",assignedTo:me.id,status:"pending",priority:"medium",allocDays:5,startDate:TODAY,dueDate:"",notes:[],blockerNote:"",createdAt:TODAY,deliveryDate:""});
  const [err,setErr]=useState("");
  const set=(k,v)=>setF(x=>({...x,[k]:v}));
  const calcDue=(s,d)=>{ if(!s||!d) return ""; const dt=new Date(s); dt.setDate(dt.getDate()+Number(d)); return dt.toISOString().slice(0,10); };
  const pL={urgent:t.urgent,high:t.high,medium:t.medium,low:t.low};
  return (
    <Modal onClose={onClose} width="min(520px,96vw)">
      <h2 style={{color:V("text"),margin:"0 0 5px",fontSize:15,paddingLeft:36}}>{t.addTask}</h2>
      <p style={{color:V("dim"),margin:"0 0 18px",fontSize:11,paddingLeft:36}}>{t.registered}: <span style={{color:V("gold")}}>{lang==="ar"?fmt(TODAY):fmtEN(TODAY)}</span></p>
      <div style={{display:"flex",flexDirection:"column",gap:11}}>
        <div>
          <label style={lbl}>{t.taskTitle} *</label>
          <input value={f.title} onChange={e=>{set("title",e.target.value);setErr("");}} style={inp} placeholder={lang==="ar"?"عنوان المهمة...":"Task title..."}/>
          {err&&<p style={{color:"#e07060",fontSize:11,margin:"3px 0 0"}}>{err}</p>}
        </div>
        <div><label style={lbl}>{t.taskTitle} (EN)</label><input value={f.titleEn} onChange={e=>set("titleEn",e.target.value)} style={inp} placeholder="Task title in English..." dir="ltr"/></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11}}>
          <div><label style={lbl}>{t.assignTo}</label><select value={f.assignedTo} onChange={e=>set("assignedTo",e.target.value)} style={sel}>{users.filter(u=>u.active).map(u=><option key={u.id} value={u.id}>{lang==="ar"?u.name:u.nameEn}</option>)}</select></div>
          <div><label style={lbl}>{t.priority}</label><select value={f.priority} onChange={e=>set("priority",e.target.value)} style={sel}>{Object.entries(pL).map(([k,v])=><option key={k} value={k}>{v}</option>)}</select></div>
          <div><label style={lbl}>{t.startD}</label><input type="date" value={f.startDate} onChange={e=>setF(x=>({...x,startDate:e.target.value,dueDate:calcDue(e.target.value,x.allocDays)}))} style={inp}/></div>
          <div><label style={lbl}>{t.allocDays}</label><input type="number" min={1} max={365} value={f.allocDays} onChange={e=>setF(x=>({...x,allocDays:e.target.value,dueDate:calcDue(x.startDate,e.target.value)}))} style={inp}/></div>
          <div style={{gridColumn:"1/-1"}}><label style={lbl}>{t.dueD} <span style={{color:V("gold"),fontSize:9,fontWeight:400}}>({t.autoCalc})</span></label><input type="date" value={f.dueDate} onChange={e=>set("dueDate",e.target.value)} style={{...inp,borderColor:f.dueDate?"var(--gold)55":V("input-border")}}/></div>
        </div>
      </div>
      <div style={{display:"flex",gap:9,marginTop:18,justifyContent:"flex-end"}}>
        <button onClick={onClose} style={mkBtnO()}>{t.cancel}</button>
        <button onClick={()=>{ if(!f.title.trim()){setErr(lang==="ar"?"يرجى كتابة عنوان المهمة":"Please enter a title");return;} onAdd({...f,id:uid()}); onClose(); }} style={mkBtnG()}>{t.addTask}</button>
      </div>
    </Modal>
  );
}

// ═══════════════════════════════════════════════════════════════
//  AI ANALYSIS MODAL
// ═══════════════════════════════════════════════════════════════
function AIModal({ tasks, users, lang, onClose }) {
  const [loading,setLoading]=useState(true);
  const [result,setResult]=useState("");
  const t=TR[lang];

  useEffect(()=>{
    const done=tasks.filter(x=>x.status==="done"&&x.deliveryDate);
    const blk=tasks.filter(x=>x.status==="blocked");
    const late=tasks.filter(x=>isLate(x));
    const workers=users.filter(u=>u.active&&u.role!=="admin_sys");
    const notes=tasks.flatMap(tk=>[...tk.notes.map(n=>n.text),tk.blockerNote]).filter(Boolean);
    const yaserN=notes.filter(n=>n.includes("ياسر")||n.includes("Yasser")).length;
    const blkReasons=tasks.filter(tk=>tk.blockerNote).map(tk=>tk.blockerNote);
    const wStats=workers.map(u=>{
      const ut=tasks.filter(tk=>tk.assignedTo===u.id);
      const ud=ut.filter(tk=>tk.status==="done"&&tk.deliveryDate);
      const effArr=ud.map(tk=>eff(tk)).filter(v=>v!=null);
      return {name:lang==="ar"?u.name:u.nameEn||u.name,total:ut.length,done:ud.length,late:ud.filter(tk=>isLate(tk)).length,onTime:ud.filter(tk=>onTime(tk)).length,blocked:ut.filter(tk=>tk.status==="blocked").length,avgEff:avg(effArr)};
    });
    const prompt=lang==="ar"
      ?`أنت محلل أداء. حلّل هذه البيانات بالعربية:\n\nمهام: ${tasks.length}، مكتملة: ${done.length}، متأخرة: ${late.length}، عوائق: ${blk.length}\n\nأداء الموظفين:\n${wStats.map(w=>`- ${w.name}: ${w.total} مهمة، ${w.done} منجزة، ${w.late} متأخرة، كفاءة ${w.avgEff!=null?w.avgEff+"%":"—"}`).join("\n")}\n\nملاحظات متكررة: ${notes.slice(0,6).join(" | ")}\nذكر "ياسر": ${yaserN} مرة\nأسباب العوائق: ${blkReasons.join(" | ")||"لا يوجد"}\n\nاكتب:\n**📊 ملخص عام**\n**✅ نقاط القوة**\n**⚠️ نقاط تحتاج تحسين**\n**🔄 أنماط متكررة**\n**💡 التوصيات** (3 نقاط)`
      :`You are a performance analyst. Analyze this data:\n\nTasks: ${tasks.length}, Done: ${done.length}, Late: ${late.length}, Blocked: ${blk.length}\n\nEmployee stats:\n${wStats.map(w=>`- ${w.name}: ${w.total} tasks, ${w.done} done, ${w.late} late, efficiency ${w.avgEff!=null?w.avgEff+"%":"N/A"}`).join("\n")}\n\nNotes: ${notes.slice(0,6).join(" | ")}\nYasser mentions: ${yaserN}\nBlockers: ${blkReasons.join(" | ")||"None"}\n\nWrite:\n**📊 Summary**\n**✅ Strengths**\n**⚠️ Areas for Improvement**\n**🔄 Recurring Patterns**\n**💡 Recommendations** (3 items)`;

    fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:900,messages:[{role:"user",content:prompt}]})})
      .then(r=>r.json()).then(d=>{ const txt=d.content?.map(c=>c.text||"").join("")||""; setResult(txt||(lang==="ar"?"تعذّر الاتصال":"Connection error")); setLoading(false); })
      .catch(()=>{
        const cr=tasks.length?Math.round(done.length/tasks.length*100):0;
        setResult(lang==="ar"
          ?`**📊 ملخص عام**\nإجمالي ${tasks.length} مهمة بنسبة إنجاز ${cr}%. ${late.length} متأخرة و${blk.length} محجوبة.\n\n**✅ نقاط القوة**\n• نظام تتبع منظم يوفر رؤية واضحة\n• بعض الموظفين يُنجزون المهام قبل الموعد\n\n**⚠️ نقاط تحتاج تحسين**\n• ${blk.length} مهمة محجوبة تحتاج تدخلاً عاجلاً\n${late.length>0?`• ${late.length} مهمة سُلِّمت بعد الموعد`:""}\n\n**🔄 أنماط متكررة**\n${yaserN>0?`• "ياسر" مذكور ${yaserN} مرات — يُوصى بمراجعة عهدته`:"• لا توجد أنماط واضحة"}\n\n**💡 التوصيات**\n1. ${yaserN>0?"مراجعة عهدة ياسر":"تحسين متابعة المستندات"}\n2. اجتماع أسبوعي لحل العوائق\n3. مراجعة توزيع الأيام المخصصة`
          :`**📊 Summary**\n${tasks.length} tasks, ${cr}% completion. ${late.length} late, ${blk.length} blocked.\n\n**✅ Strengths**\n• Structured tracking provides clear visibility\n• Several team members show strong delivery\n\n**⚠️ Areas for Improvement**\n• ${blk.length} blocked tasks need immediate action\n${late.length>0?`• ${late.length} tasks delivered past deadline`:""}\n\n**🔄 Recurring Patterns**\n${yaserN>0?`• "Yasser" appears ${yaserN} times in notes — review pending items`:"• No clear recurring patterns"}\n\n**💡 Recommendations**\n1. ${yaserN>0?"Review Yasser's pending deliverables":"Improve document tracking"}\n2. Weekly blocker review meeting\n3. Review allocated days per task`
        );
        setLoading(false);
      });
  },[]);

  return (
    <div onClick={e=>e.target===e.currentTarget&&onClose()} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:600,backdropFilter:"blur(8px)"}}>
      <div style={{background:V("modal-bg"),border:"1px solid rgba(92,184,178,.35)",borderRadius:16,width:"min(620px,96vw)",maxHeight:"88vh",overflowY:"auto",padding:30,boxShadow:V("shadow"),animation:"fadeUp .3s ease",position:"relative"}}>
        <button onClick={onClose} style={{position:"absolute",top:13,left:13,background:V("bg3"),border:`1px solid ${V("border")}`,color:V("mid"),borderRadius:7,width:29,height:29,cursor:"pointer",fontSize:12}}>✕</button>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20}}>
          <div style={{width:38,height:38,background:"linear-gradient(135deg,rgba(92,184,178,.2),rgba(200,169,126,.15))",border:"1px solid rgba(92,184,178,.35)",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🤖</div>
          <div><h2 style={{color:V("text"),margin:0,fontSize:15,fontWeight:800}}>{t.aiTitle}</h2><p style={{color:V("dim"),margin:0,fontSize:10}}>TALA Intelligence</p></div>
        </div>
        {loading
          ?<div style={{textAlign:"center",padding:"36px 0"}}><div style={{fontSize:32,marginBottom:10,display:"inline-block",animation:"spin 1.2s linear infinite"}}>⟳</div><p style={{color:V("mid"),fontSize:13}}>{t.analyzing}</p></div>
          :<div style={{fontSize:13,lineHeight:1.9}}>
            {result.split("\n").map((line,i)=>{
              if(line.startsWith("**")&&line.endsWith("**")) return <div key={i} style={{color:V("gold"),fontWeight:700,fontSize:14,marginTop:14,marginBottom:4}}>{line.replace(/\*\*/g,"")}</div>;
              if(line.startsWith("•")||line.startsWith("-")||/^\d\./.test(line)) return <div key={i} style={{color:V("text"),paddingRight:8,marginBottom:2,opacity:.85}}>{line}</div>;
              return <div key={i} style={{color:line?V("text"):V("dim"),opacity:line?.9:1}}>{line||" "}</div>;
            })}
          </div>
        }
        <div style={{display:"flex",justifyContent:"flex-end",marginTop:20}}><button onClick={onClose} style={mkBtnG()}>{t.cancel}</button></div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  EXPORT CSV  (supports optional month filter)
// ═══════════════════════════════════════════════════════════════
function doExport(tasks, users, lang, monthKey) {
  const q=s=>`"${String(s||"").replace(/"/g,'""')}"`;
  const label = monthKey||"ALL";
  const lines=[];
  lines.push(q(lang==="ar"?`=== المهام — ${label} ===`:`=== TASKS — ${label} ===`));
  lines.push([lang==="ar"?"العنوان":"Title",lang==="ar"?"المسند إليه":"Assigned To",lang==="ar"?"الحالة":"Status",lang==="ar"?"الأولوية":"Priority",lang==="ar"?"البداية":"Start",lang==="ar"?"الموعد":"Due",lang==="ar"?"التسليم":"Delivered",lang==="ar"?"أيام مخصصة":"Alloc Days",lang==="ar"?"أيام فعلية":"Actual Days",lang==="ar"?"الكفاءة":"Efficiency",lang==="ar"?"في الموعد":"On Time"].map(q).join(","));
  tasks.forEach(tk=>{
    const u=users.find(x=>x.id===tk.assignedTo);
    const al=tk.allocDays||tDur(tk),ac=aDur(tk),ev=eff(tk);
    lines.push([lang==="ar"?tk.title:tk.titleEn||tk.title,lang==="ar"?u?.name:u?.nameEn||u?.name,lang==="ar"?S_CFG[tk.status].ar:S_CFG[tk.status].en,tk.priority,tk.startDate||"",tk.dueDate||"",tk.deliveryDate||"",al||"",ac!=null?ac:"",ev!=null?ev+"%":"",onTime(tk)?(lang==="ar"?"نعم":"Yes"):isLate(tk)?(lang==="ar"?"لا":"No"):""].map(q).join(","));
  });
  lines.push("","");
  lines.push(q(lang==="ar"?`=== KPIs — ${label} ===`:`=== KPIs — ${label} ===`));
  lines.push([lang==="ar"?"الموظف":"Employee",lang==="ar"?"الدور":"Role","Total","Done","On-Time","Late","Blocked","Completion%","OnTime%","AvgAlloc","AvgActual","AvgEff"].map(q).join(","));
  users.filter(u=>u.active).forEach(u=>{
    const ut=tasks.filter(tk=>tk.assignedTo===u.id);
    const ud=ut.filter(tk=>tk.status==="done"&&tk.deliveryDate);
    const uOT=ud.filter(tk=>onTime(tk)).length;
    const effArr=ud.map(tk=>eff(tk)).filter(v=>v!=null);
    lines.push([lang==="ar"?u.name:u.nameEn||u.name,ROLE_LABEL[lang][u.role],ut.length,ud.length,uOT,ud.filter(tk=>isLate(tk)).length,ut.filter(tk=>tk.status==="blocked").length,ut.length?Math.round(ud.length/ut.length*100)+"%":"",ud.length?Math.round(uOT/ud.length*100)+"%":"",avg(ud.map(tk=>tDur(tk)).filter(v=>v!=null))||"",avg(ud.map(tk=>aDur(tk)).filter(v=>v!=null))||"",avg(effArr)?avg(effArr)+"%":""].map(q).join(","));
  });
  const blob=new Blob(["\uFEFF"+lines.join("\n")],{type:"text/csv;charset=utf-8;"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url; a.download=`TALA_${label}.csv`; a.click();
  URL.revokeObjectURL(url);
}

// ─── month helpers ────────────────────────────────────────────
const MONTH_NAMES_AR=["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"];
const MONTH_NAMES_EN=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const getYM = d => d ? d.slice(0,7) : null; // "2025-01"
const filterByMonth = (tks, ym) => !ym ? tks : tks.filter(tk => getYM(tk.deliveryDate)===ym || getYM(tk.startDate)===ym);
const filterDoneByMonth = (tks, ym) => !ym ? tks.filter(tk=>tk.status==="done"&&tk.deliveryDate) : tks.filter(tk=>tk.status==="done"&&tk.deliveryDate&&getYM(tk.deliveryDate)===ym);

// build list of all months that appear in tasks
function getAvailableMonths(tasks) {
  const set = new Set();
  tasks.forEach(tk=>{ if(tk.deliveryDate) set.add(getYM(tk.deliveryDate)); if(tk.startDate) set.add(getYM(tk.startDate)); });
  return [...set].filter(Boolean).sort();
}

// per-employee monthly series for chart (completion % per month)
function buildMonthlySeries(tasks, userId, months) {
  return months.map(ym=>{
    const ut=tasks.filter(tk=>tk.assignedTo===userId&&getYM(tk.startDate)===ym);
    const ud=tasks.filter(tk=>tk.assignedTo===userId&&tk.status==="done"&&tk.deliveryDate&&getYM(tk.deliveryDate)===ym);
    const uOT=ud.filter(tk=>onTime(tk)).length;
    const effArr=ud.map(tk=>eff(tk)).filter(v=>v!=null);
    return { ym, total:ut.length, done:ud.length, onTimeN:uOT, late:ud.filter(tk=>isLate(tk)).length, blocked:ut.filter(tk=>tk.status==="blocked").length, compPct:ut.length?Math.round(ud.length/ut.length*100):0, otPct:ud.length?Math.round(uOT/ud.length*100):0, avgEff:avg(effArr) };
  });
}

// ═══════════════════════════════════════════════════════════════
//  MONTHLY CHART (SVG line+bar)
// ═══════════════════════════════════════════════════════════════
function MonthlyChart({ series, color, lang, activeYM, onClickMonth }) {
  if(!series||series.length<1) return null;
  const W=460, H=110, PAD={t:10,b:28,l:10,r:10};
  const innerW=W-PAD.l-PAD.r, innerH=H-PAD.t-PAD.b;
  const n=series.length;
  const barW=Math.max(12, Math.min(32, (innerW/n)-6));
  const xPos=i=>PAD.l + (i/(n-1||1))*innerW;
  const yPos=v=>PAD.t + innerH - (v/100)*innerH;

  // line path for completion %
  const linePts=series.map((s,i)=>({ x:xPos(i), y:yPos(s.compPct) }));
  const linePath=linePts.map((p,i)=>`${i===0?"M":"L"}${p.x},${p.y}`).join(" ");

  // ot line
  const otPts=series.map((s,i)=>({ x:xPos(i), y:yPos(s.otPct) }));
  const otPath=otPts.map((p,i)=>`${i===0?"M":"L"}${p.x},${p.y}`).join(" ");

  return (
    <div style={{overflowX:"auto",marginBottom:4}}>
      <svg width={W} height={H} style={{display:"block",minWidth:W}}>
        {/* grid lines */}
        {[0,25,50,75,100].map(v=>(
          <line key={v} x1={PAD.l} y1={yPos(v)} x2={W-PAD.r} y2={yPos(v)}
            stroke="var(--border)" strokeWidth={v===0||v===100?1:.6} strokeDasharray={v===0||v===100?"none":"3,3"}/>
        ))}

        {/* bars for done count */}
        {series.map((s,i)=>{
          const bH=Math.max(2,(s.compPct/100)*innerH);
          const bX=xPos(i)-barW/2;
          const isActive=s.ym===activeYM;
          return (
            <g key={s.ym} onClick={()=>onClickMonth(s.ym===activeYM?null:s.ym)} style={{cursor:"pointer"}}>
              <rect x={bX} y={PAD.t+innerH-bH} width={barW} height={bH}
                fill={isActive?color:color+"66"} rx={3}
                style={{transition:"fill .2s"}}/>
              {/* active highlight */}
              {isActive&&<rect x={bX-1} y={PAD.t} width={barW+2} height={innerH} fill={color+"18"} rx={3}/>}
            </g>
          );
        })}

        {/* on-time line */}
        {series.length>1&&<path d={otPath} fill="none" stroke="#5aaa5a" strokeWidth={1.5} strokeDasharray="4,2" strokeLinecap="round" strokeLinejoin="round" opacity=".8"/>}

        {/* completion line */}
        {series.length>1&&<path d={linePath} fill="none" stroke={color} strokeWidth={2} strokeLinecap="round" strokeLinejoin="round"/>}

        {/* dots */}
        {linePts.map((p,i)=>(
          <circle key={i} cx={p.x} cy={p.y} r={series[i].ym===activeYM?5:3}
            fill={series[i].ym===activeYM?color:"var(--bg1)"} stroke={color} strokeWidth={1.5}
            onClick={()=>onClickMonth(series[i].ym===activeYM?null:series[i].ym)}
            style={{cursor:"pointer",transition:"r .2s"}}/>
        ))}

        {/* month labels */}
        {series.map((s,i)=>{
          const [y,m]=s.ym.split("-");
          const mIdx=parseInt(m,10)-1;
          const label=lang==="ar"?MONTH_NAMES_AR[mIdx]:MONTH_NAMES_EN[mIdx];
          return <text key={s.ym} x={xPos(i)} y={H-6} textAnchor="middle" fontSize={9}
            fill={s.ym===activeYM?"var(--gold)":"var(--dim)"} fontWeight={s.ym===activeYM?700:400}>{label}</text>;
        })}

        {/* legend */}
        <circle cx={PAD.l+6} cy={PAD.t+6} r={3} fill={color}/>
        <text x={PAD.l+12} y={PAD.t+9} fontSize={8} fill="var(--dim)">{lang==="ar"?"% الإنجاز":"Completion%"}</text>
        <line x1={PAD.l+70} y1={PAD.t+6} x2={PAD.l+82} y2={PAD.t+6} stroke="#5aaa5a" strokeWidth={1.5} strokeDasharray="3,2"/>
        <text x={PAD.l+86} y={PAD.t+9} fontSize={8} fill="var(--dim)">{lang==="ar"?"% في الموعد":"OnTime%"}</text>
      </svg>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  GLOBAL MONTHLY CHART (all employees on one chart)
// ═══════════════════════════════════════════════════════════════
function GlobalMonthlyChart({ tasks, workers, allMonths, lang, activeYM, onClickMonth }) {
  if(allMonths.length<1) return null;
  const W=560, H=130, PAD={t:14,b:30,l:10,r:10};
  const innerW=W-PAD.l-PAD.r, innerH=H-PAD.t-PAD.b;
  const n=allMonths.length;
  const xPos=i=>PAD.l+(i/(n-1||1))*innerW;
  const yPos=v=>PAD.t+innerH-(v/100)*innerH;

  // global completion per month
  const globalSeries=allMonths.map(ym=>{
    const ut=tasks.filter(tk=>getYM(tk.startDate)===ym);
    const ud=tasks.filter(tk=>tk.status==="done"&&tk.deliveryDate&&getYM(tk.deliveryDate)===ym);
    return { ym, compPct:ut.length?Math.round(ud.length/ut.length*100):0, otPct:ud.length?Math.round(ud.filter(tk=>onTime(tk)).length/ud.length*100):0, blocked:ut.filter(tk=>tk.status==="blocked").length };
  });

  return (
    <div style={{overflowX:"auto"}}>
      <svg width={W} height={H} style={{display:"block",minWidth:Math.min(W,360)}}>
        {[0,25,50,75,100].map(v=>(
          <line key={v} x1={PAD.l} y1={yPos(v)} x2={W-PAD.r} y2={yPos(v)} stroke="var(--border)" strokeWidth={.6} strokeDasharray="3,3"/>
        ))}
        <line x1={PAD.l} y1={yPos(0)} x2={W-PAD.r} y2={yPos(0)} stroke="var(--border2)" strokeWidth={1}/>

        {/* blocked bars */}
        {globalSeries.map((s,i)=>{
          const maxBlk=Math.max(...globalSeries.map(x=>x.blocked),1);
          const bH=Math.max(2,(s.blocked/maxBlk)*innerH*0.4);
          const bW=Math.max(8,Math.min(24,(innerW/n)-4));
          return <rect key={s.ym} x={xPos(i)-bW/2} y={PAD.t+innerH-bH} width={bW} height={bH}
            fill={s.ym===activeYM?"#e07060":"#e0706044"} rx={2}
            onClick={()=>onClickMonth(s.ym===activeYM?null:s.ym)} style={{cursor:"pointer"}}/>;
        })}

        {/* per-worker lines */}
        {workers.map(u=>{
          const series=allMonths.map(ym=>{ const ud=tasks.filter(tk=>tk.assignedTo===u.id&&tk.status==="done"&&tk.deliveryDate&&getYM(tk.deliveryDate)===ym); const ut=tasks.filter(tk=>tk.assignedTo===u.id&&getYM(tk.startDate)===ym); return ut.length?Math.round(ud.length/ut.length*100):null; });
          const pts=series.map((v,i)=>v!=null?{x:xPos(i),y:yPos(v)}:null).filter(Boolean);
          if(pts.length<1) return null;
          const path=pts.map((p,i)=>`${i===0?"M":"L"}${p.x},${p.y}`).join(" ");
          return <path key={u.id} d={path} fill="none" stroke={u.color} strokeWidth={1.5} strokeLinecap="round" strokeLinejoin="round" opacity=".75"/>;
        })}

        {/* global bold line */}
        {globalSeries.length>1&&(()=>{
          const pts=globalSeries.map((s,i)=>({x:xPos(i),y:yPos(s.compPct)}));
          const p=pts.map((pt,i)=>`${i===0?"M":"L"}${pt.x},${pt.y}`).join(" ");
          return <path d={p} fill="none" stroke="var(--gold)" strokeWidth={2.5} strokeLinecap="round" strokeLinejoin="round"/>;
        })()}

        {/* active month highlight */}
        {activeYM&&(()=>{ const idx=allMonths.indexOf(activeYM); if(idx<0) return null; return <line x1={xPos(idx)} y1={PAD.t} x2={xPos(idx)} y2={PAD.t+innerH} stroke="var(--gold)" strokeWidth={1} strokeDasharray="3,2" opacity=".6"/>; })()}

        {/* month labels */}
        {allMonths.map((ym,i)=>{
          const mIdx=parseInt(ym.split("-")[1],10)-1;
          const label=lang==="ar"?MONTH_NAMES_AR[mIdx]:MONTH_NAMES_EN[mIdx];
          return <text key={ym} x={xPos(i)} y={H-8} textAnchor="middle" fontSize={9}
            fill={ym===activeYM?"var(--gold)":"var(--dim)"} fontWeight={ym===activeYM?700:400}
            onClick={()=>onClickMonth(ym===activeYM?null:ym)} style={{cursor:"pointer"}}>{label}</text>;
        })}

        {/* legend */}
        <line x1={PAD.l} y1={PAD.t+6} x2={PAD.l+16} y2={PAD.t+6} stroke="var(--gold)" strokeWidth={2.5}/>
        <text x={PAD.l+20} y={PAD.t+9} fontSize={8} fill="var(--dim)">{lang==="ar"?"إجمالي الفريق":"Team Total"}</text>
        <rect x={PAD.l+100} y={PAD.t+2} width={8} height={8} fill="#e0706066" rx={1}/>
        <text x={PAD.l+112} y={PAD.t+9} fontSize={8} fill="var(--dim)">{lang==="ar"?"عوائق":"Blocked"}</text>
      </svg>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  KPI PAGE  (with monthly filter + charts)
// ═══════════════════════════════════════════════════════════════
function KPIPage({ tasks, users, t, lang }) {
  const [showAI,setShowAI]=useState(false);
  const allMonths=getAvailableMonths(tasks);

  // Range state: fromYM / toYM  (null = unbounded)
  const minYM=allMonths[0]||TODAY.slice(0,7);
  const maxYM=allMonths[allMonths.length-1]||TODAY.slice(0,7);
  const [fromYM,setFromYM]=useState(null);
  const [toYM,setToYM]=useState(null);

  const workers=users.filter(u=>u.active&&u.role!=="admin_sys");

  // Build year/month options from allMonths
  const years=[...new Set(allMonths.map(ym=>ym.slice(0,4)))].sort();
  const monthOpts=(lang==="ar"?MONTH_NAMES_AR:MONTH_NAMES_EN).map((n,i)=>({v:String(i+1).padStart(2,"0"),l:n}));

  const [fromY,setFromY]=useState("");
  const [fromM,setFromM]=useState("");
  const [toY,setToY]=useState("");
  const [toM,setToM]=useState("");

  // Compute active range
  const activeFrom=(fromY&&fromM)?`${fromY}-${fromM}`:null;
  const activeTo=(toY&&toM)?`${toY}-${toM}`:null;
  const hasRange=!!(activeFrom||activeTo);

  const inRange=tk=>{
    if(!hasRange) return true;
    const tkYM=getYM(tk.deliveryDate)||getYM(tk.startDate);
    if(!tkYM) return false;
    if(activeFrom&&tkYM<activeFrom) return false;
    if(activeTo&&tkYM>activeTo) return false;
    return true;
  };

  const fTasks=hasRange ? tasks.filter(tk=>inRange(tk)) : tasks;
  const done=fTasks.filter(tk=>tk.status==="done"&&tk.deliveryDate);

  const gOT=done.filter(tk=>onTime(tk)).length;
  const allEffs=done.map(tk=>eff(tk)).filter(v=>v!=null);
  const gEff=avg(allEffs), gActual=avg(done.map(tk=>aDur(tk)).filter(v=>v!=null));
  const compPct=fTasks.length?Math.round(done.length/fTasks.length*100):0;
  const otPct=done.length?Math.round(gOT/done.length*100):0;

  const rangeLabel=hasRange
    ?[activeFrom&&(()=>{const[y,m]=activeFrom.split("-");return`${lang==="ar"?MONTH_NAMES_AR[+m-1]:MONTH_NAMES_EN[+m-1]} ${y}`})(),activeTo&&(()=>{const[y,m]=activeTo.split("-");return`${lang==="ar"?MONTH_NAMES_AR[+m-1]:MONTH_NAMES_EN[+m-1]} ${y}`})()].filter(Boolean).join(lang==="ar"?" → ":" → ")
    :(lang==="ar"?"كل الوقت":"All Time");

  const clearRange=()=>{setFromY("");setFromM("");setToY("");setToM("");};

  // Chart: highlight range
  const activeYMForChart=activeFrom||activeTo||null;

  const selStyle={...sel,width:"auto",padding:"5px 9px",fontSize:12,borderRadius:7};

  return (
    <div>
      {/* ── Header ── */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{color:V("text"),margin:"0 0 3px",fontSize:16,fontWeight:800}}>{t.dashboard}</h2>
          <p style={{color:V("dim"),margin:0,fontSize:10}}>TALA KPIs</p>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <button onClick={()=>setShowAI(true)} style={{...mkBtnO(),color:V("teal"),borderColor:"var(--teal)44",fontSize:11,display:"flex",alignItems:"center",gap:5}}>{t.aiBtn}<span style={{color:V("dim"),fontSize:9}}>~10s</span></button>
          <button onClick={()=>doExport(fTasks,users,lang,rangeLabel)} style={{...mkBtnO(),color:"#5aaa5a",borderColor:"rgba(90,170,90,.3)",fontSize:11}}>{t.exportBtn}</button>
        </div>
      </div>

      {/* ── Range Filter ── */}
      <div style={{...CARD,padding:"12px 16px",marginBottom:14}}>
        <div style={{display:"flex",gap:14,alignItems:"center",flexWrap:"wrap"}}>
          {/* From */}
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <span style={{color:V("dim"),fontSize:11,whiteSpace:"nowrap"}}>{lang==="ar"?"من:":"From:"}</span>
            <select value={fromM} onChange={e=>setFromM(e.target.value)} style={selStyle}>
              <option value="">{lang==="ar"?"الشهر":"Month"}</option>
              {monthOpts.map(o=><option key={o.v} value={o.v}>{o.l}</option>)}
            </select>
            <select value={fromY} onChange={e=>setFromY(e.target.value)} style={{...selStyle,minWidth:68}}>
              <option value="">{lang==="ar"?"السنة":"Year"}</option>
              {years.map(y=><option key={y} value={y}>{y}</option>)}
            </select>
          </div>

          <span style={{color:V("border2"),fontSize:14}}>→</span>

          {/* To */}
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <span style={{color:V("dim"),fontSize:11,whiteSpace:"nowrap"}}>{lang==="ar"?"إلى:":"To:"}</span>
            <select value={toM} onChange={e=>setToM(e.target.value)} style={selStyle}>
              <option value="">{lang==="ar"?"الشهر":"Month"}</option>
              {monthOpts.map(o=><option key={o.v} value={o.v}>{o.l}</option>)}
            </select>
            <select value={toY} onChange={e=>setToY(e.target.value)} style={{...selStyle,minWidth:68}}>
              <option value="">{lang==="ar"?"السنة":"Year"}</option>
              {years.map(y=><option key={y} value={y}>{y}</option>)}
            </select>
          </div>

          {/* Active label + clear */}
          <div style={{display:"flex",alignItems:"center",gap:8,marginRight:"auto"}}>
            <span style={{color:hasRange?V("gold"):V("dim"),fontSize:12,fontWeight:hasRange?700:400}}>{rangeLabel}</span>
            {hasRange&&<button onClick={clearRange} style={{...mkBtnO(),fontSize:10,padding:"2px 9px",color:"#e07060",borderColor:"rgba(224,112,96,.3)"}}>✕ {lang==="ar"?"إعادة ضبط":"Reset"}</button>}
          </div>
        </div>
      </div>

      {/* ── Global chart ── */}
      {allMonths.length>1&&(
        <div style={{...CARD,padding:"14px 16px",marginBottom:14}}>
          <p style={{color:V("dim"),fontSize:9,margin:"0 0 10px",letterSpacing:".06em",textTransform:"uppercase"}}>
            {lang==="ar"?"الأداء الشهري للفريق":"Monthly Team Performance"}
          </p>
          <GlobalMonthlyChart tasks={tasks} workers={workers} allMonths={allMonths} lang={lang} activeYM={activeYMForChart} onClickMonth={()=>{}}/>
          {/* range highlight overlay hint */}
          {hasRange&&<p style={{color:V("dim"),fontSize:9,margin:"6px 0 0",textAlign:"center"}}>{lang==="ar"?"الفترة المختارة مظللة في الرسم":"Selected range highlighted in chart"}</p>}
        </div>
      )}

      {/* ── Global KPI circles ── */}
      <div style={{...CARD,padding:"16px 20px",marginBottom:14}}>
        <div style={{display:"flex",gap:14,alignItems:"center",justifyContent:"space-around",flexWrap:"wrap"}}>
          <CircleStat pct={compPct} color="var(--gold)" size={70} label={compPct+"%"} sub={t.compRate}/>
          <CircleStat pct={otPct} color="#5aaa5a" size={70} label={otPct+"%"} sub={t.otRate}/>
          {gEff!=null&&<CircleStat pct={Math.min(100,gEff)} color={gEff>=100?"#5aaa5a":"#e07060"} size={70} label={gEff+"%"} sub={t.avgEff}/>}
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {[[t.total,fTasks.length,"#4a9eba"],[t.done,done.length,"#5aaa5a"],[t.onTimeL,gOT,"#5aaa5a"],[t.lateL,done.filter(tk=>isLate(tk)).length,"#e07060"]].map(([l,v,c])=>(
              <div key={l} style={{display:"flex",justifyContent:"space-between",gap:16,alignItems:"center"}}>
                <span style={{color:V("dim"),fontSize:11}}>{l}</span><span style={{color:c,fontSize:13,fontWeight:800}}>{v}</span>
              </div>
            ))}
          </div>
          {gActual!=null&&<div style={{textAlign:"center"}}><p style={{color:V("dim"),fontSize:9,margin:"0 0 4px"}}>{t.avgActual}</p><p style={{color:V("gold"),fontSize:22,fontWeight:800,margin:0}}>{gActual}<span style={{fontSize:11,fontWeight:400}}> {t.dayL}</span></p></div>}
        </div>
      </div>

      {/* ── Per-employee ── */}
      <h3 style={{color:V("text"),margin:"0 0 11px",fontSize:13,fontWeight:700}}>{t.empPerf}</h3>
      <div style={{display:"flex",flexDirection:"column",gap:11}}>
        {workers.map(u=>{
          const uFilt=fTasks.filter(tk=>tk.assignedTo===u.id);
          const uDone=done.filter(tk=>tk.assignedTo===u.id);
          const uOT=uDone.filter(tk=>onTime(tk)).length, uLate=uDone.filter(tk=>isLate(tk)).length;
          const uBlk=uFilt.filter(tk=>tk.status==="blocked").length, uOD=uFilt.filter(tk=>isOD(tk)).length;
          const effArr=uDone.map(tk=>eff(tk)).filter(v=>v!=null);
          const uEff=avg(effArr), uActual=avg(uDone.map(tk=>aDur(tk)).filter(v=>v!=null)), uAlloc=avg(uDone.map(tk=>tDur(tk)).filter(v=>v!=null));
          const compP=uFilt.length?Math.round(uDone.length/uFilt.length*100):0;
          const otP=uDone.length?Math.round(uOT/uDone.length*100):0;
          const series=buildMonthlySeries(tasks,u.id,allMonths);

          return (
            <div key={u.id} style={{...CARD,padding:16}}>
              {/* identity + circles + badges */}
              <div style={{display:"flex",gap:14,alignItems:"center",marginBottom:14,flexWrap:"wrap"}}>
                <div style={{display:"flex",alignItems:"center",gap:9,minWidth:130}}>
                  <div style={{width:42,height:42,borderRadius:"50%",background:V("bg0"),border:`2px solid ${u.color}55`,color:u.color,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:800,flexShrink:0}}>{u.avatar}</div>
                  <div>
                    <p style={{color:V("text"),margin:0,fontSize:13,fontWeight:700}}>{lang==="ar"?u.name:u.nameEn||u.name}</p>
                    <p style={{color:ROLE_COLOR[u.role],margin:0,fontSize:9}}>{ROLE_LABEL[lang][u.role]}</p>
                  </div>
                </div>
                <div style={{display:"flex",gap:11,alignItems:"center"}}>
                  <CircleStat pct={compP} color={u.color} size={56} label={compP+"%"} sub={t.compRate}/>
                  <CircleStat pct={otP} color={otP>=80?"#5aaa5a":otP>=50?"#c8a97e":"#e07060"} size={56} label={otP+"%"} sub={t.otRate}/>
                  {uEff!=null&&<CircleStat pct={Math.min(100,uEff)} color={uEff>=100?"#5aaa5a":"#e07060"} size={56} label={uEff+"%"} sub={t.avgEff}/>}
                </div>
                <div style={{display:"flex",gap:6,flexWrap:"wrap",marginRight:"auto"}}>
                  {[[t.total,uFilt.length,V("mid")],[t.done,uDone.length,"#5aaa5a"],[t.onTimeL,uOT,"#4a9eba"],[t.lateL,uLate,"#e07060"],[t.blocked,uBlk,"#c8a97e"],[t.overdue,uOD,"#e07060"]].map(([l,v,c])=>(
                    <div key={l} style={{background:V("badge-bg"),borderRadius:7,padding:"4px 9px",textAlign:"center",minWidth:48}}>
                      <p style={{color:c,margin:0,fontSize:14,fontWeight:800}}>{v}</p>
                      <p style={{color:V("dim"),margin:0,fontSize:9}}>{l}</p>
                    </div>
                  ))}
                </div>
              </div>

              {/* timing KPIs */}
              {uDone.length>0&&(
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12}}>
                  {[[t.avgAlloc,uAlloc!=null?`${uAlloc} ${t.dayL}`:"—","var(--gold)"],[t.avgActual,uActual!=null?`${uActual} ${t.dayL}`:"—",uActual&&uAlloc&&uActual<=uAlloc?"#5aaa5a":"#e07060"],[t.avgEff,uEff!=null?`${uEff}%`:"—",uEff>=100?"#5aaa5a":uEff>=80?"#c8a97e":"#e07060"]].map(([l,v,c])=>(
                    <div key={l} style={{background:V("bg3"),border:`1px solid ${V("border")}`,borderRadius:8,padding:"8px 12px",flex:1,minWidth:110}}>
                      <p style={{color:V("dim"),fontSize:9,margin:"0 0 3px"}}>{l}</p>
                      <p style={{color:c,fontSize:16,fontWeight:800,margin:0}}>{v}</p>
                    </div>
                  ))}
                </div>
              )}

              {/* monthly chart per employee */}
              {allMonths.length>1&&series.some(s=>s.total>0)&&(
                <div style={{background:V("bg0"),borderRadius:8,padding:"10px 12px",marginBottom:10}}>
                  <p style={{color:V("dim"),fontSize:9,margin:"0 0 6px",letterSpacing:".06em",textTransform:"uppercase"}}>{lang==="ar"?"التطور الشهري":"Monthly Progress"}</p>
                  <MonthlyChart series={series.filter(s=>s.total>0||s.done>0)} color={u.color} lang={lang} activeYM={activeYMForChart} onClickMonth={()=>{}}/>
                  {/* monthly data mini-cards */}
                  <div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:8}}>
                    {series.filter(s=>s.total>0||s.done>0).map(s=>{ const[y,m]=s.ym.split("-"); const ml=lang==="ar"?MONTH_NAMES_AR[+m-1]:MONTH_NAMES_EN[+m-1]; const inSel=hasRange?((!activeFrom||s.ym>=activeFrom)&&(!activeTo||s.ym<=activeTo)):true; return (
                      <div key={s.ym} style={{background:inSel?`${u.color}18`:V("bg3"),border:`1px solid ${inSel?u.color+"44":V("border")}`,borderRadius:8,padding:"6px 10px",minWidth:72,textAlign:"center",transition:"all .2s"}}>
                        <p style={{color:inSel?u.color:V("mid"),fontSize:10,fontWeight:700,margin:"0 0 3px"}}>{ml} {y.slice(2)}</p>
                        <p style={{color:"#5aaa5a",fontSize:12,fontWeight:800,margin:"0 0 1px"}}>{s.done}/{s.total}</p>
                        <p style={{color:s.compPct>=80?"#5aaa5a":s.compPct>=50?"#c8a97e":"#e07060",fontSize:10,fontWeight:700,margin:0}}>{s.compPct}%</p>
                        {s.blocked>0&&<p style={{color:"#e07060",fontSize:9,margin:"2px 0 0"}}>⚠ {s.blocked}</p>}
                      </div>);}
                    )}
                  </div>
                </div>
              )}

              {/* task details */}
              {uDone.length>0&&(
                <div>
                  <p style={{color:V("dim"),fontSize:9,margin:"0 0 5px",letterSpacing:".06em",textTransform:"uppercase"}}>{t.detailsL}</p>
                  {uDone.map(tk=>{ const al=tDur(tk),ac=aDur(tk),ev=eff(tk),ot=onTime(tk); return (
                    <div key={tk.id} style={{background:V("bg0"),borderRadius:7,padding:"6px 10px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:6,marginBottom:3}}>
                      <span style={{color:V("text"),fontSize:11,flex:1}}>{lang==="ar"?tk.title:tk.titleEn||tk.title}</span>
                      <div style={{display:"flex",gap:8}}>
                        {al&&<span style={{color:V("dim"),fontSize:10}}>{lang==="ar"?"مخصص":"alloc"}: {al}{lang==="ar"?"د":"d"}</span>}
                        {ac!=null&&<span style={{color:ac<=al?"#5aaa5a":"#e07060",fontSize:10}}>{lang==="ar"?"فعلي":"actual"}: {ac}{lang==="ar"?"د":"d"}</span>}
                        {ev!=null&&<span style={{color:ev>=100?"#5aaa5a":"#e07060",fontSize:10,fontWeight:700}}>{ev}%</span>}
                        <span style={{color:ot?"#5aaa5a":"#e07060",fontSize:10}}>{ot?(lang==="ar"?"✓ في الموعد":"✓ on-time"):(lang==="ar"?"⚡ متأخر":"⚡ late")}</span>
                      </div>
                    </div>);}
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
      {showAI&&<AIModal tasks={tasks} users={users} lang={lang} onClose={()=>setShowAI(false)}/>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  PERMISSIONS PAGE
// ═══════════════════════════════════════════════════════════════
function PermsPage({ rolePerms, setRolePerms, t, lang }) {
  const PK=["canAdd","canAssign","canViewAll","canManageUsers","canReports","canApprove","canEditRoles"];
  const ROLES=["admin_sys","dept_mgr","senior","junior"];
  const toggle=(role,pk)=>{ if(role==="admin_sys") return; setRolePerms(p=>({...p,[role]:{...p[role],[pk]:!p[role][pk]}})); };
  return (
    <div>
      <div style={{marginBottom:20}}>
        <h2 style={{color:V("text"),margin:"0 0 4px",fontSize:16,fontWeight:800}}>{t.permTitle}</h2>
        <p style={{color:V("dim"),margin:0,fontSize:12}}>{t.permDesc}</p>
      </div>
      <div style={{...CARD,overflow:"hidden"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:V("bg1"),borderBottom:`1px solid ${V("border")}`}}>
            <th style={{padding:"12px 15px",color:V("mid"),fontWeight:600,textAlign:lang==="ar"?"right":"left",fontSize:10}}>{t.role}</th>
            {PK.map(pk=><th key={pk} style={{padding:"12px 9px",color:V("mid"),fontWeight:600,textAlign:"center",fontSize:10,whiteSpace:"nowrap"}}>{t[pk]}</th>)}
          </tr></thead>
          <tbody>{ROLES.map((role,ri)=>(
            <tr key={role} style={{borderBottom:ri<ROLES.length-1?`1px solid ${V("border")}`:"none",background:ri%2===0?"transparent":V("table-alt")}}>
              <td style={{padding:"11px 15px"}}>
                <div style={{display:"flex",alignItems:"center",gap:7}}>
                  <div style={{width:7,height:7,borderRadius:"50%",background:ROLE_COLOR[role]}}/>
                  <span style={{color:V("text"),fontWeight:600,fontSize:12}}>{ROLE_LABEL[lang][role]}</span>
                </div>
              </td>
              {PK.map(pk=>(
                <td key={pk} style={{padding:"11px 9px",textAlign:"center"}}>
                  <button onClick={()=>toggle(role,pk)} style={{width:27,height:27,borderRadius:7,border:`1px solid ${rolePerms[role]?.[pk]?"rgba(90,170,90,.4)":V("border")}`,background:rolePerms[role]?.[pk]?"rgba(90,170,90,.15)":V("bg0"),cursor:role==="admin_sys"?"not-allowed":"pointer",color:rolePerms[role]?.[pk]?"#5aaa5a":V("dim"),fontSize:13,display:"inline-flex",alignItems:"center",justifyContent:"center",opacity:role==="admin_sys"?.4:1,transition:"all .2s"}}>{rolePerms[role]?.[pk]?"✓":"–"}</button>
                </td>
              ))}
            </tr>
          ))}</tbody>
        </table>
      </div>
      <p style={{color:V("dim"),fontSize:10,margin:"8px 0 0",textAlign:"center"}}>{lang==="ar"?"مدير النظام دائماً لديه كل الصلاحيات":"System Admin always retains all permissions"}</p>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  USERS PAGE
// ═══════════════════════════════════════════════════════════════
function UsersPage({ users, setUsers, tasks, t, lang }) {
  const [editing,setEditing]=useState(null); const [ef,setEf]=useState({});
  const [showAdd,setShowAdd]=useState(false);
  const [nf,setNf]=useState({name:"",nameEn:"",email:"",password:"",role:"junior",active:true,color:"#5aaa5a"});
  const COLORS=["#c8a97e","#4a9eba","#5aaa5a","#9a6aba","#e07060","#4abaa0","#d4a04a"];
  const add=()=>{ if(!nf.name||!nf.email||!nf.password)return; const av=nf.name.split(" ").map(w=>w[0]).join("").slice(0,2); setUsers(us=>[...us,{...nf,id:uid(),avatar:av}]); setNf({name:"",nameEn:"",email:"",password:"",role:"junior",active:true,color:"#5aaa5a"}); setShowAdd(false); };
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <h2 style={{color:V("text"),margin:0,fontSize:16,fontWeight:800}}>👥 {t.users}</h2>
        <button onClick={()=>setShowAdd(s=>!s)} style={mkBtnG()}>{t.addUser}</button>
      </div>
      {showAdd&&(
        <div style={{...CARD,padding:17,marginBottom:17,border:"1px solid rgba(200,169,126,.35)"}}>
          <h3 style={{color:V("gold"),margin:"0 0 13px",fontSize:13}}>{t.newUser}</h3>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><label style={lbl}>{t.fullname} (AR)</label><input value={nf.name} onChange={e=>setNf(n=>({...n,name:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>{t.fullname} (EN)</label><input value={nf.nameEn} onChange={e=>setNf(n=>({...n,nameEn:e.target.value}))} style={inp} dir="ltr"/></div>
            <div><label style={lbl}>{t.email}</label><input value={nf.email} onChange={e=>setNf(n=>({...n,email:e.target.value}))} style={inp} dir="ltr"/></div>
            <div><label style={lbl}>{t.pass}</label><input type="password" value={nf.password} onChange={e=>setNf(n=>({...n,password:e.target.value}))} style={inp}/></div>
            <div><label style={lbl}>{t.role}</label><select value={nf.role} onChange={e=>setNf(n=>({...n,role:e.target.value}))} style={sel}>{Object.keys(ROLE_LABEL.ar).map(k=><option key={k} value={k}>{ROLE_LABEL[lang][k]}</option>)}</select></div>
            <div><label style={lbl}>{t.colorL}</label><div style={{display:"flex",gap:7,marginTop:4}}>{COLORS.map(c=><div key={c} onClick={()=>setNf(n=>({...n,color:c}))} style={{width:25,height:25,borderRadius:"50%",background:c,cursor:"pointer",border:nf.color===c?"3px solid var(--text)":"2px solid transparent",transition:"all .15s"}}/>)}</div></div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:13,justifyContent:"flex-end"}}><button onClick={()=>setShowAdd(false)} style={mkBtnO()}>{t.cancel}</button><button onClick={add} style={mkBtnG()}>{t.addUser}</button></div>
        </div>
      )}
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {users.map(u=>{ const tc=tasks.filter(tk=>tk.assignedTo===u.id).length; const dc=tasks.filter(tk=>tk.assignedTo===u.id&&tk.status==="done").length;
          return (
          <div key={u.id} style={{...CARD,padding:"13px 17px",display:"flex",alignItems:"center",gap:11,opacity:u.active?1:.45,transition:"opacity .2s"}}>
            <div style={{width:40,height:40,borderRadius:"50%",background:V("bg0"),border:`2px solid ${u.color}55`,color:u.color,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:13,flexShrink:0}}>{u.avatar}</div>
            {editing===u.id
              ?<div style={{flex:1,display:"grid",gridTemplateColumns:"1fr 1fr 1fr auto",gap:9,alignItems:"end"}}>
                <div><label style={lbl}>{t.name}</label><input value={ef.name} onChange={e=>setEf(f=>({...f,name:e.target.value}))} style={inp}/></div>
                <div><label style={lbl}>{t.email}</label><input value={ef.email} onChange={e=>setEf(f=>({...f,email:e.target.value}))} style={inp} dir="ltr"/></div>
                <div><label style={lbl}>{t.role}</label><select value={ef.role} onChange={e=>setEf(f=>({...f,role:e.target.value}))} style={sel}>{Object.keys(ROLE_LABEL.ar).map(k=><option key={k} value={k}>{ROLE_LABEL[lang][k]}</option>)}</select></div>
                <div style={{display:"flex",gap:7}}><button onClick={()=>{setUsers(us=>us.map(x=>x.id===editing?{...x,...ef}:x));setEditing(null);}} style={mkBtnG()}>{t.save}</button><button onClick={()=>setEditing(null)} style={mkBtnO()}>×</button></div>
              </div>
              :<div style={{flex:1,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
                <div><p style={{color:V("text"),margin:"0 0 2px",fontSize:13,fontWeight:600}}>{lang==="ar"?u.name:u.nameEn||u.name}</p><p style={{color:V("dim"),margin:0,fontSize:11}} dir="ltr">{u.email}</p></div>
                <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                  <span style={{color:ROLE_COLOR[u.role],fontSize:11,background:`${ROLE_COLOR[u.role]}20`,border:`1px solid ${ROLE_COLOR[u.role]}44`,borderRadius:20,padding:"2px 9px"}}>{ROLE_LABEL[lang][u.role]}</span>
                  <span style={{color:V("mid"),fontSize:11}}>📋 {tc} · ✅ {dc}</span>
                  <div style={{display:"flex",gap:6}}>
                    <button onClick={()=>{setEditing(u.id);setEf({...u});}} style={{...mkBtnO(),padding:"4px 10px",fontSize:11}}>{t.editU}</button>
                    {u.role!=="admin_sys"&&<button onClick={()=>setUsers(us=>us.map(x=>x.id===u.id?{...x,active:!x.active}:x))} style={{...mkBtnO(),padding:"4px 10px",fontSize:11,color:u.active?"#e07060":"#5aaa5a",borderColor:u.active?"rgba(224,112,96,.3)":"rgba(90,170,90,.3)"}}>{u.active?t.disableU:t.enableU}</button>}
                  </div>
                </div>
              </div>
            }
          </div>);
        })}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  TASK CARD
// ═══════════════════════════════════════════════════════════════
function TaskCard({ task, users, onClick, lang }) {
  const asn=users.find(u=>u.id===task.assignedTo);
  const s=S_CFG[task.status]; const od=isOD(task);
  const [hov,setHov]=useState(false);
  const al=task.allocDays||tDur(task); const ac=aDur(task);

  return (
    <div onClick={onClick} onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)}
      style={{background:hov?V("hover-bg"):V("card-bg"),border:`1px solid ${od?"rgba(224,112,96,.4)":hov?V("border2"):V("card-border")}`,borderRight:lang==="ar"?`3px solid ${s.color}`:"none",borderLeft:lang==="en"?`3px solid ${s.color}`:"none",borderRadius:10,padding:"13px 14px",cursor:"pointer",transition:"all .2s",transform:hov?"translateY(-2px)":"none",boxShadow:hov?V("shadow"):"none"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:7,gap:7}}>
        <div style={{display:"flex",gap:5,alignItems:"flex-start",flex:1}}>
          <span style={{width:5,height:5,borderRadius:"50%",background:PC[task.priority],flexShrink:0,marginTop:6}}/>
          <h3 style={{color:V("text"),fontSize:12,fontWeight:600,margin:0,lineHeight:1.6}}>{lang==="ar"?task.title:task.titleEn||task.title}</h3>
        </div>
        <span style={{color:s.color,background:dark_?s.dbg:s.dlg,border:`1px solid ${s.color}44`,borderRadius:20,padding:"1px 8px",fontSize:10,whiteSpace:"nowrap",flexShrink:0}}>{s.icon} {lang==="ar"?s.ar:s.en}</span>
      </div>
      {task.status==="blocked"&&task.blockerNote&&<div style={{background:"rgba(224,112,96,.1)",border:"1px solid rgba(224,112,96,.3)",borderRadius:7,padding:"4px 9px",marginBottom:8}}><p style={{color:"#e07060",fontSize:10,margin:0}}>⚠ {task.blockerNote}</p></div>}
      {al&&<div style={{marginBottom:8}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}><span style={{color:V("dim"),fontSize:9}}>{ac!=null?`${ac} / ${al} ${lang==="ar"?"يوم":"d"}`:`${al} ${lang==="ar"?"يوم مخصص":"days alloc."}`}</span>{ac!=null&&<span style={{color:ac<=al?"#5aaa5a":"#e07060",fontSize:9,fontWeight:700}}>{eff({startDate:task.startDate,dueDate:task.dueDate,deliveryDate:task.deliveryDate})}%</span>}</div><Bar pct={ac!=null?(ac/al)*100:0} color={ac!=null&&ac<=al?"#5aaa5a":"#e07060"}/></div>}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:7}}>
        <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
          {task.startDate&&<span style={{color:V("dim"),fontSize:10}}>▶ {lang==="ar"?fmt(task.startDate):fmtEN(task.startDate)}</span>}
          {task.dueDate&&<span style={{color:od?"#e07060":V("dim"),fontSize:10}}>⏰ {lang==="ar"?fmt(task.dueDate):fmtEN(task.dueDate)}{od?" ⚡":""}</span>}
          {task.deliveryDate&&<span style={{color:"#5aaa5a",fontSize:10}}>✓ {lang==="ar"?fmt(task.deliveryDate):fmtEN(task.deliveryDate)}</span>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:5}}>
          {task.notes.length>0&&<span style={{color:V("dim"),fontSize:10}}>💬 {task.notes.length}</span>}
          <div style={{width:22,height:22,borderRadius:"50%",background:V("bg0"),border:`1px solid ${asn?.color||"var(--gold)"}55`,color:asn?.color||"var(--gold)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700}}>{asn?.avatar}</div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  CORRECTION PANEL
// ═══════════════════════════════════════════════════════════════
function CorrPanel({ requests, users, onApprove, onReject, t, lang }) {
  const pending=requests.filter(r=>r.status==="pending");
  if(!pending.length) return null;
  return (
    <div style={{...CARD,border:"1px solid rgba(200,169,126,.35)",marginBottom:18,overflow:"hidden"}}>
      <div style={{background:"rgba(200,169,126,.08)",borderBottom:"1px solid rgba(200,169,126,.25)",padding:"9px 15px",display:"flex",alignItems:"center",gap:7}}>
        <span>🔔</span><span style={{color:V("gold"),fontWeight:700,fontSize:13}}>{t.corrPanel} ({pending.length})</span>
      </div>
      {pending.map((r,i)=>{ const u=users.find(x=>x.id===r.requestedBy); return (
        <div key={r.id} style={{padding:"11px 15px",borderBottom:i<pending.length-1?`1px solid ${V("border")}`:"none",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
          <div>
            <p style={{color:V("text"),margin:"0 0 2px",fontSize:12,fontWeight:600}}>{r.taskTitle}</p>
            <p style={{color:V("mid"),margin:"0 0 2px",fontSize:11}}>{t.reqBy}: <span style={{color:u?.color}}>{lang==="ar"?u?.name:u?.nameEn||u?.name}</span> — {t.reqDate}: <span style={{color:V("gold"),fontWeight:600}}>{lang==="ar"?fmt(r.requestedDate):fmtEN(r.requestedDate)}</span></p>
            <p style={{color:V("dim"),margin:0,fontSize:10}}>{t.reasonL}: {r.reason}</p>
          </div>
          <div style={{display:"flex",gap:6}}>
            <button onClick={()=>onApprove(r)} style={{...mkBtnO(),color:"#5aaa5a",borderColor:"rgba(90,170,90,.3)",fontSize:11,padding:"4px 11px"}}>{t.approveBtn}</button>
            <button onClick={()=>onReject(r.id)} style={{...mkBtnO(),color:"#e07060",borderColor:"rgba(224,112,96,.3)",fontSize:11,padding:"4px 11px"}}>{t.rejectBtn}</button>
          </div>
        </div>);
      })}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  APP SHELL
// ═══════════════════════════════════════════════════════════════
function AppShell({ me, users, setUsers, tasks, setTasks, notifs, setNotifs, corrReqs, setCorrReqs, rolePerms, setRolePerms, lang, setLang, dark, setDark, onLogout }) {
  const [page,setPage]=useState("dashboard");
  const [sel,setSel]=useState(null);
  const [showAdd,setShowAdd]=useState(false);
  const [fStatus,setFS]=useState("all");
  const [fUser,setFU]=useState("all");
  const [view,setView]=useState("board");

  const t=TR[lang];
  const dir=lang==="ar"?"rtl":"ltr";
  const perms=rolePerms[me.role]||SEED_PERMS[me.role]||SEED_PERMS.junior;
  const myTasks=perms.canViewAll?tasks:tasks.filter(tk=>tk.assignedTo===me.id);
  const visible=myTasks.filter(tk=>(fStatus==="all"||tk.status===fStatus)&&(!perms.canViewAll||fUser==="all"||tk.assignedTo===fUser));
  const stats={total:myTasks.length,done:myTasks.filter(tk=>tk.status==="done").length,prog:myTasks.filter(tk=>tk.status==="inprogress").length,blocked:myTasks.filter(tk=>tk.status==="blocked").length,late:myTasks.filter(tk=>isOD(tk)).length};
  const myNotifs=notifs.filter(n=>n.toUser===me.id).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
  const pendingCorr=corrReqs.filter(r=>r.status==="pending").length;

  const push=(toUser,type,text)=>setNotifs(ns=>[...ns,{id:uid(),toUser,type,text,read:false,createdAt:new Date().toISOString()}]);
  const handleComplete=(taskId,dDate)=>{ setTasks(ts=>ts.map(tk=>tk.id===taskId?{...tk,status:"done",deliveryDate:dDate}:tk)); users.filter(u=>u.role==="admin_sys"||u.role==="dept_mgr").forEach(a=>{ const tk=tasks.find(x=>x.id===taskId); push(a.id,"task_completed",lang==="ar"?`✅ أنجز ${me.name} مهمة: "${tk?.title}"`:`✅ ${me.nameEn||me.name} completed: "${tk?.titleEn||tk?.title}"`); }); };
  const handleAdd=(task)=>{ setTasks(ts=>[...ts,task]); if(task.assignedTo!==me.id) push(task.assignedTo,"task_assigned",lang==="ar"?`📋 تم تعيين مهمة لك: "${task.title}"`:`📋 New task assigned: "${task.titleEn||task.title}"`); };
  const handleSave=(updated)=>{ const old=tasks.find(tk=>tk.id===updated.id); setTasks(ts=>ts.map(tk=>tk.id===updated.id?updated:tk)); if(old&&old.assignedTo!==updated.assignedTo) push(updated.assignedTo,"task_assigned",lang==="ar"?`📋 تم تعيين مهمة لك: "${updated.title}"`:`📋 Task assigned: "${updated.titleEn||updated.title}"`); if(old&&old.status!=="blocked"&&updated.status==="blocked") users.filter(u=>u.role==="admin_sys"||u.role==="dept_mgr").forEach(a=>push(a.id,"task_blocked",lang==="ar"?`⚠ عائق: "${updated.title}"`:`⚠ Blocked: "${updated.titleEn||updated.title}"`)); };
  const handleCorrReq=(req)=>{ setCorrReqs(rs=>[...rs,req]); users.filter(u=>u.role==="admin_sys"||u.role==="dept_mgr").forEach(a=>push(a.id,"correction_requested",lang==="ar"?`🔔 طلب تصحيح: "${req.taskTitle}" من ${me.name}`:`🔔 Correction request from ${me.nameEn||me.name}`)); };
  const approveCo=(req)=>{ setTasks(ts=>ts.map(tk=>tk.id===req.taskId?{...tk,deliveryDate:req.requestedDate}:tk)); setCorrReqs(rs=>rs.map(r=>r.id===req.id?{...r,status:"approved"}:r)); push(req.requestedBy,"correction_approved",lang==="ar"?`✓ وافق المدير على تصحيح تاريخ: "${req.taskTitle}"`:`✓ Correction approved: "${req.taskTitle}"`); };
  const rejectCo=(id)=>{ const req=corrReqs.find(r=>r.id===id); setCorrReqs(rs=>rs.map(r=>r.id===id?{...r,status:"rejected"}:r)); if(req) push(req.requestedBy,"correction_rejected",lang==="ar"?`✕ رفض المدير التصحيح: "${req.taskTitle}"`:`✕ Correction rejected: "${req.taskTitle}"`); };

  const NAV=[
    {id:"dashboard",icon:"◈",label:t.dashboard},
    {id:"tasks",icon:"▦",label:perms.canViewAll?"جميع المهام":t.myTasks},
    ...(perms.canManageUsers?[{id:"users",icon:"⊕",label:t.users}]:[]),
    ...(perms.canEditRoles?[{id:"perms",icon:"⊗",label:t.perms}]:[]),
  ];

  return (
    <div dir={dir} style={{minHeight:"100vh",background:V("bg0"),display:"flex"}}>
      {/* SIDEBAR */}
      <aside style={{width:214,background:V("sidebar-bg"),borderLeft:lang==="ar"?`1px solid ${V("border")}`:"none",borderRight:lang==="en"?`1px solid ${V("border")}`:"none",display:"flex",flexDirection:"column",position:"fixed",top:0,[lang==="ar"?"right":"left"]:0,bottom:0,zIndex:50,transition:"background .3s"}}>
        <div style={{padding:"18px 16px 14px",borderBottom:`1px solid ${V("border")}`}}>
          <div style={{display:"flex",alignItems:"center",gap:9}}>
            <div style={{width:34,height:34,borderRadius:9,background:dark?"linear-gradient(135deg,#1e1810,#161214)":"linear-gradient(135deg,#faf5ec,#f0e8d8)",border:"1px solid rgba(200,169,126,.45)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
              <svg viewBox="0 0 24 24" fill="none" style={{width:20,height:20}}>
                <path d="M4 20 L12 4 L20 20" stroke="var(--gold)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M7 15 L17 15" stroke="var(--gold)" strokeWidth="1.5" strokeLinecap="round" opacity=".5"/>
                <circle cx="12" cy="21.5" r="1.5" fill="var(--teal)"/>
              </svg>
            </div>
            <div>
              <p style={{color:V("gold"),margin:0,fontSize:16,fontWeight:800,letterSpacing:"1.5px"}}>{t.appName}</p>
              <p style={{color:V("dim"),margin:0,fontSize:11,fontWeight:700,letterSpacing:".02em"}}>{lang==="ar"?"نظام تتبع المهام":"Task Management"}</p>
            </div>
          </div>
        </div>

        <div style={{padding:"12px 15px",borderBottom:`1px solid ${V("border")}`}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:34,height:34,borderRadius:"50%",background:V("bg0"),border:`2px solid ${me.color}55`,color:me.color,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800,flexShrink:0}}>{me.avatar}</div>
            <div style={{overflow:"hidden",flex:1}}>
              <p style={{color:V("text"),margin:"0 0 1px",fontSize:12,fontWeight:600,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{lang==="ar"?me.name:me.nameEn||me.name}</p>
              <p style={{color:ROLE_COLOR[me.role],margin:0,fontSize:9}}>{ROLE_LABEL[lang][me.role]}</p>
            </div>
          </div>
        </div>

        <nav style={{flex:1,padding:"9px 7px",display:"flex",flexDirection:"column",gap:2}}>
          {NAV.map(n=>(
            <button key={n.id} onClick={()=>setPage(n.id)}
              style={{background:page===n.id?"rgba(200,169,126,.1)":"none",border:`1px solid ${page===n.id?"rgba(200,169,126,.3)":"transparent"}`,color:page===n.id?V("gold"):V("mid"),borderRadius:9,padding:"9px 12px",cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:page===n.id?700:400,display:"flex",alignItems:"center",gap:9,transition:"all .2s",position:"relative"}}
              onMouseEnter={e=>{ if(page!==n.id) e.currentTarget.style.background=V("hover-bg"); }}
              onMouseLeave={e=>{ if(page!==n.id) e.currentTarget.style.background="none"; }}>
              <span style={{fontSize:13,opacity:.8}}>{n.icon}</span>
              <span style={{flex:1,textAlign:lang==="ar"?"right":"left"}}>{n.label}</span>
              {n.id==="tasks"&&pendingCorr>0&&perms.canApprove&&<span style={{background:"#e07060",color:"white",borderRadius:"50%",width:16,height:16,fontSize:9,display:"inline-flex",alignItems:"center",justifyContent:"center",fontWeight:800}}>{pendingCorr}</span>}
            </button>
          ))}
        </nav>

        <div style={{padding:"11px 13px",borderTop:`1px solid ${V("border")}`}}>
          {[[t.done,stats.done,"#5aaa5a"],[t.prog,stats.prog,"#4a9eba"],[t.overdue,stats.late,"#e07060"]].map(([l,v,c])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{color:V("dim"),fontSize:10}}>{l}</span><span style={{color:c,fontSize:11,fontWeight:700}}>{v}</span></div>
          ))}
        </div>

        <div style={{padding:"7px 7px 11px"}}>
          <button onClick={onLogout} style={{...mkBtnO(),width:"100%",textAlign:"center",color:"#e07060",borderColor:"rgba(224,112,96,.25)",fontSize:11}}>
            {lang==="ar"?"⬅ ":""}{t.logout}{lang==="en"?" →":""}
          </button>
        </div>
      </aside>

      {/* MAIN */}
      <main style={{flex:1,[lang==="ar"?"marginRight":"marginLeft"]:214,minHeight:"100vh",display:"flex",flexDirection:"column",transition:"background .3s"}}>
        {/* TOPBAR */}
        <div style={{background:V("topbar-bg"),borderBottom:`1px solid ${V("border")}`,padding:"0 22px",height:50,display:"flex",justifyContent:"space-between",alignItems:"center",position:"sticky",top:0,zIndex:40,transition:"background .3s"}}>
          <h1 style={{color:V("text"),margin:0,fontSize:14,fontWeight:700}}>{NAV.find(n=>n.id===page)?.label||""}</h1>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <button onClick={()=>setLang(l=>l==="ar"?"en":"ar")} style={{...mkBtnO(),fontSize:10,padding:"3px 8px",color:V("gold"),borderColor:"var(--gold)33"}}>{lang==="ar"?"EN":"ع"}</button>
            <ThemeBtn dark={dark} onToggle={()=>setDark(d=>!d)} t={t}/>
            <NotifBell notifs={myNotifs} onRead={id=>setNotifs(ns=>ns.map(n=>n.id===id?{...n,read:true}:n))} onReadAll={()=>setNotifs(ns=>ns.map(n=>n.toUser===me.id?{...n,read:true}:n))} t={t}/>
            {perms.canAdd&&page==="tasks"&&<button onClick={()=>setShowAdd(true)} style={mkBtnG()}>{t.addTask}</button>}
          </div>
        </div>

        {/* CONTENT */}
        <div style={{padding:"20px 22px",flex:1}}>
          {page==="dashboard"&&<KPIPage tasks={tasks} users={users} t={t} lang={lang}/>}

          {page==="tasks"&&<>
            {perms.canApprove&&<CorrPanel requests={corrReqs} users={users} onApprove={approveCo} onReject={rejectCo} t={t} lang={lang}/>}
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:9,marginBottom:16}}>
              {[[t.total,stats.total,"#4a9eba","▦"],[t.done,stats.done,"#5aaa5a","✓"],[t.prog,stats.prog,"#4a9eba","⟳"],[stats.late>0?t.overdue:t.blocked,stats.late>0?stats.late:stats.blocked,"#e07060",stats.late>0?"⚡":"⚠"]].map(([l,v,c,ic])=>(
                <div key={l} style={{...CARD,padding:"13px 15px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div><p style={{color:V("mid"),margin:"0 0 2px",fontSize:10}}>{l}</p><p style={{color:c,margin:0,fontSize:23,fontWeight:800}}>{v}</p></div>
                  <span style={{fontSize:18,opacity:.4,color:V("text")}}>{ic}</span>
                </div>
              ))}
            </div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:13,flexWrap:"wrap",gap:8}}>
              <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                {[["all",t.allStatus],...Object.entries(S_CFG).map(([k,v])=>[k,`${v.icon} ${lang==="ar"?v.ar:v.en}`])].map(([k,l])=>(
                  <button key={k} onClick={()=>setFS(k)} style={{background:fStatus===k?V("bg3"):"none",border:`1px solid ${fStatus===k?V("border2"):V("border")}`,color:fStatus===k?(S_CFG[k]?.color||V("gold")):V("dim"),borderRadius:20,padding:"4px 12px",cursor:"pointer",fontSize:11,fontFamily:"inherit",transition:"all .2s"}}>{l}</button>
                ))}
                {perms.canViewAll&&<select value={fUser} onChange={e=>setFU(e.target.value)} style={{...sel,width:"auto",borderRadius:20,fontSize:11,padding:"4px 10px"}}>
                  <option value="all">{t.allEmp}</option>
                  {users.filter(u=>u.active).map(u=><option key={u.id} value={u.id}>{lang==="ar"?u.name:u.nameEn||u.name}</option>)}
                </select>}
              </div>
              <div style={{display:"flex",gap:5}}>
                {[[t.boardV,"board"],[t.listV,"list"]].map(([l,v])=><button key={v} onClick={()=>setView(v)} style={{background:view===v?V("bg3"):"none",border:`1px solid ${view===v?V("border2"):V("border")}`,color:view===v?V("gold"):V("dim"),borderRadius:7,padding:"4px 11px",cursor:"pointer",fontSize:11,transition:"all .2s"}}>{l}</button>)}
              </div>
            </div>
            {view==="board"
              ?<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(255px,1fr))",gap:10}}>
                {visible.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:"50px 0",color:V("dim")}}><p style={{fontSize:30,margin:"0 0 8px"}}>📭</p><p style={{fontSize:13}}>{t.noTasks}</p></div>}
                {visible.map(tk=><TaskCard key={tk.id} task={tk} users={users} onClick={()=>setSel(tk)} lang={lang}/>)}
              </div>
              :<div style={{...CARD,overflow:"hidden"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{background:V("bg1"),borderBottom:`1px solid ${V("border")}`}}>
                    {[t.tasks,...(perms.canViewAll?[t.assignTo]:[]),t.status,t.startD,t.dueD,t.delivD,t.allocDays,"Eff.","💬"].map(h=><th key={h} style={{padding:"9px 11px",color:V("mid"),fontWeight:600,textAlign:lang==="ar"?"right":"left",fontSize:10,whiteSpace:"nowrap"}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {visible.length===0&&<tr><td colSpan={9} style={{textAlign:"center",padding:34,color:V("dim")}}>{t.noTasks}</td></tr>}
                    {visible.map((tk,i)=>{ const s=S_CFG[tk.status]; const u=users.find(x=>x.id===tk.assignedTo); const od=isOD(tk); const al=tk.allocDays||tDur(tk); const ac=aDur(tk); const ev=eff(tk);
                      return (<tr key={tk.id} onClick={()=>setSel(tk)} style={{borderBottom:`1px solid ${V("border")}`,cursor:"pointer",background:i%2===0?"transparent":V("table-alt"),transition:"background .15s"}}
                        onMouseEnter={e=>e.currentTarget.style.background=V("hover-bg")}
                        onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"transparent":V("table-alt")}>
                        <td style={{padding:"9px 11px",color:V("text"),maxWidth:180}}>{lang==="ar"?tk.title:tk.titleEn||tk.title}</td>
                        {perms.canViewAll&&<td style={{padding:"9px 11px"}}><span style={{display:"flex",alignItems:"center",gap:5}}><span style={{width:20,height:20,borderRadius:"50%",background:V("bg0"),border:`1px solid ${u?.color}44`,color:u?.color,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:8,fontWeight:700}}>{u?.avatar}</span><span style={{color:V("mid"),fontSize:11}}>{lang==="ar"?u?.name:u?.nameEn||u?.name}</span></span></td>}
                        <td style={{padding:"9px 11px"}}><span style={{color:s.color,background:dark_?s.dbg:s.dlg,border:`1px solid ${s.color}44`,borderRadius:20,padding:"1px 7px",fontSize:10}}>{s.icon} {lang==="ar"?s.ar:s.en}</span></td>
                        <td style={{padding:"9px 11px",color:V("dim"),fontSize:11}}>{lang==="ar"?fmt(tk.startDate):fmtEN(tk.startDate)}</td>
                        <td style={{padding:"9px 11px",color:od?"#e07060":V("dim"),fontSize:11}}>{lang==="ar"?fmt(tk.dueDate):fmtEN(tk.dueDate)}</td>
                        <td style={{padding:"9px 11px",color:"#5aaa5a",fontSize:11}}>{lang==="ar"?fmt(tk.deliveryDate):fmtEN(tk.deliveryDate)}</td>
                        <td style={{padding:"9px 11px",fontSize:11}}>{al&&<span style={{color:V("mid")}}>{al}</span>}{ac!=null&&<span style={{color:ac<=al?"#5aaa5a":"#e07060"}}> / {ac}</span>}</td>
                        <td style={{padding:"9px 11px",fontSize:11,fontWeight:600,color:ev!=null?(ev>=100?"#5aaa5a":"#e07060"):V("dim")}}>{ev!=null?ev+"%":"—"}</td>
                        <td style={{padding:"9px 11px",color:V("dim"),fontSize:11}}>{tk.notes.length||"—"}</td>
                      </tr>);
                    })}
                  </tbody>
                </table>
              </div>
            }
          </>}

          {page==="users"&&perms.canManageUsers&&<UsersPage users={users} setUsers={setUsers} tasks={tasks} t={t} lang={lang}/>}
          {page==="perms"&&perms.canEditRoles&&<PermsPage rolePerms={rolePerms} setRolePerms={setRolePerms} t={t} lang={lang}/>}
        </div>
      </main>

      {sel&&<TaskModal task={sel} users={users} me={me} perms={perms} onClose={()=>setSel(null)} onSave={handleSave} onDelete={id=>setTasks(ts=>ts.filter(tk=>tk.id!==id))} onComplete={handleComplete} onCorrReq={handleCorrReq} t={t} lang={lang}/>}
      {showAdd&&perms.canAdd&&<AddTaskModal users={users} me={me} onClose={()=>setShowAdd(false)} onAdd={handleAdd} t={t} lang={lang}/>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  ROOT
// ═══════════════════════════════════════════════════════════════
export default function App() {
  const [users,     setUsers]    = useState(()=>sg("tala_users",    SEED_USERS));
  const [tasks,     setTasks]    = useState(()=>sg("tala_tasks",    SEED_TASKS));
  const [notifs,    setNotifs]   = useState(()=>sg("tala_notifs",   []));
  const [corrReqs,  setCorrReqs] = useState(()=>sg("tala_corr",     []));
  const [rolePerms, setRolePerms]= useState(()=>sg("tala_perms",    SEED_PERMS));
  const [lang,      setLang]     = useState(()=>sg("tala_lang",     "ar"));
  const [dark,      setDark]     = useState(()=>sg("tala_dark",     true));
  const [me,        setMe]       = useState(null);

  // update global dark_ marker for components that can't receive props easily
  dark_ = dark;

  useEffect(()=>ss("tala_users",    users),    [users]);
  useEffect(()=>ss("tala_tasks",    tasks),    [tasks]);
  useEffect(()=>ss("tala_notifs",   notifs),   [notifs]);
  useEffect(()=>ss("tala_corr",     corrReqs), [corrReqs]);
  useEffect(()=>ss("tala_perms",    rolePerms),[rolePerms]);
  useEffect(()=>ss("tala_lang",     lang),     [lang]);
  useEffect(()=>ss("tala_dark",     dark),     [dark]);

  const font = lang==="ar"
    ? "'Sakkal Majalla','Noto Naskh Arabic','Traditional Arabic',serif"
    : "'Playfair Display','Times New Roman',Georgia,serif";
  const dir  = lang==="ar" ? "rtl" : "ltr";
  const cls  = dark ? "tala-dark" : "tala-light";

  return (
    <div className={cls} dir={dir} style={{fontFamily:font,minHeight:"100vh",background:V("bg0"),color:V("text"),transition:"background .3s, color .3s"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;500;600;700&family=Playfair+Display:ital,wght@0,400;0,500;0,600;0,700;0,800;1,400&display=swap');
  /* Arabic: Noto Naskh Arabic (web-safe Sakkal Majalla fallback) */
  /* English: Playfair Display (Times New Roman feel, modern serif) */
        ${THEME_CSS}
      `}</style>
      {!me
        ?<Login users={users} onLogin={u=>setMe(users.find(x=>x.id===u.id)||u)} lang={lang} setLang={setLang} dark={dark} setDark={setDark}/>
        :<AppShell me={me} users={users} setUsers={setUsers} tasks={tasks} setTasks={setTasks} notifs={notifs} setNotifs={setNotifs} corrReqs={corrReqs} setCorrReqs={setCorrReqs} rolePerms={rolePerms} setRolePerms={setRolePerms} lang={lang} setLang={setLang} dark={dark} setDark={setDark} onLogout={()=>setMe(null)}/>
      }
    </div>
  );
}
