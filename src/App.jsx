import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════════
// TRANSLATIONS
// ═══════════════════════════════════════════════════════════════
const T = {
  ar: {
    dir:"rtl", font:"'Cairo',sans-serif",
    appName:"TALA COSTING",
    login:"تسجيل الدخول", userId:"رقم المستخدم", pin:"PIN (5 أرقام)", loginBtn:"دخول", loginErr:"رقم المستخدم أو PIN غير صحيح",
    dashboard:"لوحة التحليل", rawMat:"المواد الخام", prepItem:"المواد شبه المصنعة", products:"المنتجات",
    classes:"إدارة الكلاسات", users:"إدارة المستخدمين", logout:"خروج",
    code:"الكود", name:"الاسم", class:"الكلاس", unit:"الوحدة", price:"السعر", actions:"إجراءات",
    kg:"كيلو", liter:"لتر", piece:"حبة",
    add:"إضافة", edit:"تعديل", delete:"حذف", save:"حفظ", cancel:"إلغاء", search:"بحث...",
    exportXlsx:"تصدير", importXlsx:"استيراد",
    importNote:"يحدث فقط المواد الموجودة — بالاسم أولاً ثم بالكود",
    noData:"لا توجد بيانات", confirmDelete:"هل تريد الحذف؟",
    ingredients:"المكونات", addIngredient:"إضافة مكون", ingredient:"المادة", qty:"الكمية", waste:"هدر %",
    yieldWeight:"وزن الناتج", yieldAuto:"(تلقائي)", costPerUnit:"التكلفة/وحدة",
    totalCost:"إجمالي التكلفة", sellingPrice:"سعر البيع", margin:"هامش %",
    rawMatCategory:"Raw Material", prepCategory:"Prep Item",
    className:"اسم الكلاس", addClass:"إضافة كلاس", classFor:"الكلاسات لـ",
    required:"هذا الحقل مطلوب", positiveNum:"يجب أن يكون أكبر من صفر",
    savedOk:"تم الحفظ", deletedOk:"تم الحذف", importedOk:"تم تحديث", importedItems:"صنف", noMatch:"لم يتم العثور على تطابق",
    source:"المصدر", productName:"اسم المنتج", all:"الكل", lang:"EN",
    // users
    userName:"اسم المستخدم", userRole:"الدور", admin:"مدير", user:"مستخدم",
    permissions:"الصلاحيات", permView:"عرض", permEdit:"تعديل", permDelete:"حذف",
    addUser:"إضافة مستخدم", editUser:"تعديل مستخدم", newPin:"PIN جديد",
    pinHint:"5 أرقام فقط", userIdLabel:"رقم المستخدم",
    // dashboard
    show:"عرض", items:"صنف",
    secProducts:"تحليل المنتجات", secPrep:"تحليل المواد شبه المصنعة", secRaw:"تحليل المواد الخام",
    highPriceLowCost:"أعلى سعر + أقل تكلفة", highPriceHighCost:"أعلى سعر + أعلى تكلفة",
    highCostLowPrice:"أعلى تكلفة + أقل سعر",
    prepHighCost:"أعلى تكلفة/وحدة", prepMostUsed:"الأكثر استخداماً في المنتجات",
    rawHighPrice:"أعلى سعر", rawMostUsed:"الأكثر استخداماً",
    usedIn:"مستخدم في", productsCount:"منتج",
    noPermission:"ليس لديك صلاحية للوصول لهذا القسم",
    // standard costing
    stdCost:"التكلفة المعيارية", variance:"الانحراف", variancePct:"نسبة الانحراف",
    aboveStd:"فوق المعياري", belowStd:"تحت المعياري", onTarget:"مطابق",
    topCostDriver:"أعلى مكون تكلفة",
    // usage columns
    usedInPrep:"في Prep", usedInProducts:"في المنتجات",
    clickToView:"اضغط للعرض",
    relatedProducts:"المنتجات المرتبطة", relatedPreps:"Prep المرتبطة",
    // dashboard sections
    kpiSection:"المؤشرات الرئيسية",
    varianceReport:"تقرير الانحراف عن المعياري",
    marginDistribution:"توزيع هامش الربح",
    topCostDriversReport:"أعلى مكونات التكلفة",
    costBreakdown:"تفصيل التكلفة",
    aboveTarget:"تجاوز المعياري", belowTarget:"دون المعياري", noTarget:"بدون معياري",
    totalProducts:"إجمالي المنتجات", totalPrep:"إجمالي Prep", totalRaw:"إجمالي الخام",
    avgMarginLbl:"متوسط الهامش", avgCostLbl:"متوسط التكلفة",
    highMargin:"هامش عالي >30%", midMargin:"هامش متوسط 15-30%", lowMargin:"هامش منخفض <15%",
    noSelling:"بدون سعر بيع",
    productsOverBudget:"منتجات تجاوزت المعياري",
    view:"عرض",
  },
  en: {
    dir:"ltr", font:"'DM Sans',sans-serif",
    appName:"TALA COSTING",
    login:"Login", userId:"User ID", pin:"PIN (5 digits)", loginBtn:"Sign In", loginErr:"Invalid User ID or PIN",
    dashboard:"Analytics", rawMat:"Raw Materials", prepItem:"Prep Items", products:"Products",
    classes:"Manage Classes", users:"Manage Users", logout:"Logout",
    code:"Code", name:"Name", class:"Class", unit:"Unit", price:"Price", actions:"Actions",
    kg:"Kg", liter:"Liter", piece:"Piece",
    add:"Add", edit:"Edit", delete:"Delete", save:"Save", cancel:"Cancel", search:"Search...",
    exportXlsx:"Export", importXlsx:"Import",
    importNote:"Updates existing items only — by name first, then by code",
    noData:"No data found", confirmDelete:"Delete this item?",
    ingredients:"Ingredients", addIngredient:"Add Ingredient", ingredient:"Material", qty:"Qty", waste:"Waste %",
    yieldWeight:"Yield", yieldAuto:"(auto)", costPerUnit:"Cost/Unit",
    totalCost:"Total Cost", sellingPrice:"Selling Price", margin:"Margin %",
    rawMatCategory:"Raw Material", prepCategory:"Prep Item",
    className:"Class Name", addClass:"Add Class", classFor:"Classes for",
    required:"This field is required", positiveNum:"Must be greater than zero",
    savedOk:"Saved", deletedOk:"Deleted", importedOk:"Updated", importedItems:"items", noMatch:"No matching items found",
    source:"Source", productName:"Product Name", all:"All", lang:"ع",
    userName:"Username", userRole:"Role", admin:"Admin", user:"User",
    permissions:"Permissions", permView:"View", permEdit:"Edit", permDelete:"Delete",
    addUser:"Add User", editUser:"Edit User", newPin:"New PIN",
    pinHint:"5 digits only", userIdLabel:"User ID",
    show:"Show", items:"Items",
    secProducts:"Products Analysis", secPrep:"Prep Items Analysis", secRaw:"Raw Materials Analysis",
    highPriceLowCost:"High Price + Low Cost", highPriceHighCost:"High Price + High Cost",
    highCostLowPrice:"High Cost + Low Price",
    prepHighCost:"Highest Cost/Unit", prepMostUsed:"Most Used in Products",
    rawHighPrice:"Highest Price", rawMostUsed:"Most Used",
    usedIn:"Used in", productsCount:"products",
    noPermission:"You don't have permission to access this section",
    stdCost:"Std Cost", variance:"Variance", variancePct:"Var %",
    aboveStd:"Above Std", belowStd:"Below Std", onTarget:"On Target",
    topCostDriver:"Top Cost Driver",
    usedInPrep:"In Prep", usedInProducts:"In Products",
    clickToView:"Click to view",
    relatedProducts:"Related Products", relatedPreps:"Related Preps",
    kpiSection:"Key Indicators",
    varianceReport:"Variance vs Standard",
    marginDistribution:"Margin Distribution",
    topCostDriversReport:"Top Cost Drivers",
    costBreakdown:"Cost Breakdown",
    aboveTarget:"Above Target", belowTarget:"Below Target", noTarget:"No Target",
    totalProducts:"Total Products", totalPrep:"Total Prep", totalRaw:"Total Raw",
    avgMarginLbl:"Avg Margin", avgCostLbl:"Avg Cost",
    highMargin:"High Margin >30%", midMargin:"Mid Margin 15-30%", lowMargin:"Low Margin <15%",
    noSelling:"No Selling Price",
    productsOverBudget:"Products Over Budget",
    view:"View",
  }
};

// ═══════════════════════════════════════════════════════════════
// STORAGE + AUTH
// ═══════════════════════════════════════════════════════════════
const SK = { raw:"tc_raw_v1", prep:"tc_prep_v1", products:"tc_prods_v1", classes:"tc_cls_v1", users:"tc_users_v1", session:"tc_session_v1" };
const load = (k,d) => { try { const v=localStorage.getItem(k); return v?JSON.parse(v):d; } catch { return d; } };
const persist = (k,v) => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} };

const defaultUsers = [
  { id:"1001", pin:"11111", name:"Admin", role:"admin", lang:"ar",
    perms:{ raw:{view:true,edit:true,delete:true}, prep:{view:true,edit:true,delete:true}, products:{view:true,edit:true,delete:true}, classes:{view:true,edit:true,delete:true} } }
];
const defaultClasses = { raw:["Food Item","Package Item","Cleaning Item"], prep:["Sauce","Dough","Mix","Marinade"] };

const genCode = (prefix,existing) => {
  const nums = existing.map(i=>parseInt(i.code?.replace(prefix+"-",""))||0);
  const next = nums.length ? Math.max(...nums)+1 : 1;
  return `${prefix}-${String(next).padStart(5,"0")}`;
};
const unitLbl = (u,t) => ({kg:t.kg,liter:t.liter,piece:t.piece}[u]||u);

// ═══════════════════════════════════════════════════════════════
// COLORS
// ═══════════════════════════════════════════════════════════════
const DARK = {
  bg:"#080b14", sidebar:"#0d1020", surface:"#111627", card:"#161c2e",
  border:"#1e2540", accent:"#e8a020", accentDark:"#c4861a",
  text:"#dde3f0", muted:"#5a6585", danger:"#ef4444",
  green:"#22c55e", red:"#ef4444", yellow:"#f59e0b", blue:"#3b82f6",
};
const LIGHT = {
  bg:"#f4f6fb", sidebar:"#ffffff", surface:"#eef1f8", card:"#ffffff",
  border:"#dde2ef", accent:"#c4861a", accentDark:"#a06a0e",
  text:"#1a2140", muted:"#7a869e", danger:"#ef4444",
  green:"#16a34a", red:"#dc2626", yellow:"#d97706", blue:"#2563eb",
};
const C = DARK;

// ═══════════════════════════════════════════════════════════════
// APP ROOT
// ═══════════════════════════════════════════════════════════════
export default function App() {
  const [users,setUsers]     = useState(()=>load(SK.users,defaultUsers));
  const [session,setSession] = useState(()=>load(SK.session,null));
  const [rawList,setRawList] = useState(()=>load(SK.raw,[]));
  const [prepList,setPrepList]= useState(()=>load(SK.prep,[]));
  const [prodList,setProdList]= useState(()=>load(SK.products,[]));
  const [classes,setClasses] = useState(()=>load(SK.classes,defaultClasses));
  const [tab,setTab]         = useState("dashboard");
  const [toast,setToast]     = useState(null);
  const [sideOpen,setSideOpen]= useState(true);

  useEffect(()=>persist(SK.users,users),[users]);
  useEffect(()=>persist(SK.raw,rawList),[rawList]);
  useEffect(()=>persist(SK.prep,prepList),[prepList]);
  useEffect(()=>persist(SK.products,prodList),[prodList]);
  useEffect(()=>persist(SK.classes,classes),[classes]);
  useEffect(()=>persist(SK.session,session),[session]);

  const showToast = useCallback((msg,type="success")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),3000); },[]);

  const currentUser = session ? users.find(u=>u.id===session.id) : null;
  const lang = currentUser?.lang || "ar";
  const t = T[lang];

  const calcPrepCostFn = (prep, depth=0)=>{
    if(!prep.ingredients?.length||depth>5) return {totalCost:0,yieldKg:0,costPerUnit:0};
    let totalCost=0, yieldG=0;
    prep.ingredients.forEach(ing=>{
      const qty=parseFloat(ing.qty)||0;
      const waste=(parseFloat(ing.waste)||0)/100;
      const netQty=qty*(1-waste); // qty after waste (goes into yield)
      if(ing.source==="prep"){
        // prep inside prep
        const subPrep=prepList.find(p=>String(p.id)===String(ing.rawId||ing.srcId||ing.prepId));
        if(!subPrep) return;
        const sub=calcPrepCostFn(subPrep,depth+1);
        // cost = qty × costPerUnit (buy raw qty, pay for it)
        totalCost+=(subPrep.unit==="piece"?qty:qty/1000)*sub.costPerUnit;
      } else {
        const raw=rawList.find(r=>String(r.id)===String(ing.rawId));
        if(!raw) return;
        // cost = raw qty × price (no waste adjustment on cost)
        totalCost+=(raw.unit==="piece"?qty:qty/1000)*raw.price;
      }
      // yield = net qty after waste
      yieldG+=netQty;
    });
    const yieldKg=prep.yieldOverride&&parseFloat(prep.yieldOverride)>0
      ?parseFloat(prep.yieldOverride)
      :(yieldG/1000);
    return {totalCost, yieldKg, costPerUnit:yieldKg>0?totalCost/yieldKg:0};
  };
  const calcPrepCost = useCallback((prep)=>calcPrepCostFn(prep),[rawList,prepList]);

  const calcProductCost = useCallback((prod)=>{
    if(!prod.ingredients?.length) return {totalCost:0,margin:0};
    let cost=0;
    prod.ingredients.forEach(ing=>{
      const qty2=parseFloat(ing.qty)||0;
      const waste2=(parseFloat(ing.waste)||0)/100;
      if(ing.source==="raw"){
        const raw=rawList.find(r=>String(r.id)===String(ing.srcId)); if(!raw) return;
        // cost on raw qty (before waste)
        cost+=(raw.unit==="piece"?qty2:qty2/1000)*raw.price;
      } else {
        const prep=prepList.find(p=>String(p.id)===String(ing.srcId)); if(!prep) return;
        const {costPerUnit}=calcPrepCost(prep);
        cost+=(prep.unit==="piece"?qty2:qty2/1000)*costPerUnit;
      }
    });
    const sp=parseFloat(prod.sellingPrice)||0;
    return {totalCost:cost, margin:sp>0?((sp-cost)/sp)*100:0};
  },[rawList,prepList,calcPrepCost]);

  const hasPerm = (mod,action) => {
    if(!currentUser) return false;
    if(currentUser.role==="admin") return true;
    return currentUser.perms?.[mod]?.[action]===true;
  };

  
  if(!currentUser) return <LoginScreen users={users} onLogin={u=>setSession({id:u.id})} lang={lang} />;

  const navItems = [
    {id:"dashboard", label:t.dashboard},
    {id:"raw",       label:t.rawMat,    perm:"raw"},
    {id:"prep",      label:t.prepItem,  perm:"prep"},
    {id:"products",  label:t.products,  perm:"products"},
    {id:"classes",   label:t.classes,   perm:"classes"},
    ...(currentUser.role==="admin"?[{id:"users",label:t.users}]:[]),
  ];

  return (
    <div dir={t.dir} style={{display:"flex",minHeight:"100vh",background:DARK.bg,fontFamily:t.font,color:DARK.text}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;900&family=DM+Sans:wght@400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:4px;height:4px}
        ::-webkit-scrollbar-track{background:${DARK.sidebar}}
        ::-webkit-scrollbar-thumb{background:${DARK.accent}66;border-radius:4px}
        .btn{border:none;cursor:pointer;border-radius:7px;font-family:inherit;font-weight:600;transition:all .16s;display:inline-flex;align-items:center;gap:5px}
        .btn:active{transform:scale(.97)}
        .btn-primary{background:linear-gradient(135deg,${DARK.accent},${DARK.accentDark});color:#080b14;padding:8px 18px;font-size:13px}
        .btn-primary:hover{filter:brightness(1.1);box-shadow:0 4px 16px ${DARK.accent}44}
        .btn-secondary{background:${DARK.card};color:${DARK.muted};padding:8px 14px;font-size:13px;border:1px solid ${DARK.border}}
        .btn-secondary:hover{color:${DARK.text};border-color:#3a4060}
        .btn-sm-e{background:#0f2a4a;color:#60a5fa;padding:4px 11px;font-size:12px;border:1px solid #1e3a6033;border-radius:6px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
        .btn-sm-e:hover{background:#1e3a6033}
        .btn-sm-d{background:#2a0f0f;color:#f87171;padding:4px 11px;font-size:12px;border:1px solid #dc262633;border-radius:6px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
        .btn-sm-d:hover{background:#dc262622}
        input,select{background:${DARK.surface};border:1px solid ${DARK.border};color:${DARK.text};border-radius:7px;padding:8px 12px;font-family:inherit;font-size:13px;width:100%;outline:none;transition:border .16s}
        input:focus,select:focus{border-color:${DARK.accent};box-shadow:0 0 0 3px ${DARK.accent}15}
        input::placeholder{color:#2e3a55}
        input[type=number]{-moz-appearance:textfield}
        input[type=checkbox]{width:auto;cursor:pointer;accent-color:${DARK.accent}}
        .lbl{font-size:11px;color:${DARK.muted};margin-bottom:4px;display:block;font-weight:600;text-transform:uppercase;letter-spacing:.04em}
        .err{color:#f87171;font-size:11px;margin-top:3px}
        .card{background:${DARK.card};border:1px solid ${DARK.border};border-radius:12px}
        .overlay{position:fixed;inset:0;background:rgba(0,0,0,.82);display:flex;align-items:center;justify-content:center;z-index:200;backdrop-filter:blur(8px);padding:16px}
        .modal{background:${DARK.card};border:1px solid ${DARK.border};border-radius:16px;padding:24px;width:100%;max-width:520px;max-height:90vh;overflow-y:auto}
        .modal-lg{max-width:720px}
        .toast{position:fixed;bottom:18px;${lang==="ar"?"right":"left"}:18px;padding:10px 16px;border-radius:9px;font-size:13px;font-weight:600;z-index:300;animation:su .22s ease;pointer-events:none}
        .toast.success{background:#053d2a;color:#4ade80;border:1px solid #16a34a}
        .toast.error{background:#3d0505;color:#fca5a5;border:1px solid #dc2626}
        .toast.warning{background:#3d2205;color:#fcd34d;border:1px solid #ca8a04}
        @keyframes su{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
        th{padding:10px 13px;text-align:${lang==="ar"?"right":"left"};font-size:10px;font-weight:700;color:${DARK.muted};text-transform:uppercase;letter-spacing:.06em;background:${DARK.surface};white-space:nowrap}
        td{padding:10px 13px;border-bottom:1px solid ${DARK.border}22;vertical-align:middle;font-size:13px}
        tr:last-child td{border-bottom:none}
        tr:hover td{background:${DARK.surface}77}
        .badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700}
        .badge-kg{background:#0f2040;color:#60a5fa}
        .badge-liter{background:#062018;color:#4ade80}
        .badge-piece{background:#1a0f33;color:#a78bfa}
        .badge-cls{background:#1a1a33;color:#94a3b8;border:1px solid #2a2a4a}
        .filter-btn{background:${DARK.surface};border:1px solid ${DARK.border};color:${DARK.muted};padding:5px 12px;font-size:12px;border-radius:20px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
        .filter-btn.active{background:${DARK.accent}20;border-color:${DARK.accent}66;color:${DARK.accent}}
        .divider{height:1px;background:${DARK.border};margin:14px 0}
        .nav-item{display:flex;align-items:center;padding:9px 14px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;color:${DARK.muted};transition:all .14s;border:none;background:transparent;font-family:inherit;width:100%;text-align:${lang==="ar"?"right":"left"}}
        .nav-item:hover{background:${DARK.surface};color:${DARK.text}}
        .nav-item.active{background:${DARK.accent}18;color:${DARK.accent};border-${lang==="ar"?"right":"left"}:3px solid ${DARK.accent}}
        .perm-box{display:flex;gap:12px;align-items:center;padding:10px 14px;background:${DARK.surface};border-radius:8px;border:1px solid ${DARK.border};margin-bottom:6px}
        .bar-bg{background:${DARK.surface};border-radius:3px;height:6px;flex:1}
        .bar-fill{height:6px;border-radius:3px;transition:width .4s ease}
        .section-hd{font-size:14px;font-weight:700;color:${DARK.text};margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid ${DARK.border}}
        .top-n-btn{background:${DARK.surface};border:1px solid ${DARK.border};color:${DARK.muted};padding:5px 12px;font-size:12px;border-radius:6px;cursor:pointer;font-family:inherit;font-weight:700;transition:all .14s}
        .top-n-btn.active{background:${DARK.accent}20;border-color:${DARK.accent}66;color:${DARK.accent}}
      `}</style>

      {/* SIDEBAR */}
      <div style={{width:sideOpen?220:0,minWidth:sideOpen?220:0,background:DARK.sidebar,borderRight:`1px solid ${DARK.border}`,display:"flex",flexDirection:"column",transition:"all .2s",overflow:"hidden",position:"sticky",top:0,height:"100vh",flexShrink:0}}>
        <div style={{padding:"20px 16px 12px"}}>
          <div style={{fontWeight:900,fontSize:16,color:C.accent,letterSpacing:"1px"}}>{t.appName}</div>
          <div style={{fontSize:10,color:C.muted,marginTop:2,letterSpacing:".5px"}}>SYSTEM v1.0</div>
        </div>
        <div style={{height:1,background:C.border,margin:"0 12px 10px"}}/>
        <div style={{flex:1,padding:"0 8px",overflowY:"auto"}}>
          {navItems.map(n=>(
            <button key={n.id} className={`nav-item${tab===n.id?" active":""}`} onClick={()=>setTab(n.id)}>{n.label}</button>
          ))}
        </div>
        <div style={{padding:"12px 8px",borderTop:`1px solid ${DARK.border}`}}>
          <div style={{padding:"8px 14px",fontSize:12,color:C.muted,marginBottom:6}}>
            <div style={{fontWeight:700,color:C.text,fontSize:13}}>{currentUser.name}</div>
            <div style={{fontSize:11}}>{currentUser.role==="admin"?t.admin:t.user}</div>
          </div>
          <button className="nav-item" style={{color:C.danger}} onClick={()=>{ persist(SK.session,null); setSession(null); }}>{t.logout}</button>
        </div>
      </div>

      {/* MAIN */}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        {/* TOPBAR */}
        <div style={{background:DARK.sidebar,borderBottom:`1px solid ${DARK.border}`,padding:"10px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:50}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button onClick={()=>setSideOpen(o=>!o)} style={{background:"transparent",border:"none",cursor:"pointer",color:C.muted,fontSize:18,padding:"2px 6px",borderRadius:5,transition:"color .14s"}} onMouseOver={e=>e.currentTarget.style.color=C.text} onMouseOut={e=>e.currentTarget.style.color=C.muted}>☰</button>
            <span style={{fontWeight:700,fontSize:14,color:C.text}}>{navItems.find(n=>n.id===tab)?.label||""}</span>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <button className="btn btn-secondary" style={{padding:"6px 13px",fontSize:12}} onClick={()=>{ const u=users.find(u=>u.id===currentUser.id); if(u){ const newLang=u.lang==="ar"?"en":"ar"; setUsers(prev=>prev.map(x=>x.id===u.id?{...x,lang:newLang}:x)); } }}>
              {t.lang}
            </button>
          </div>
        </div>

        {/* CONTENT */}
        <div style={{flex:1,overflowY:"auto",padding:"20px"}}>
          {tab==="dashboard" && <DashboardTab t={t} lang={lang} C={DARK} rawList={rawList} prepList={prepList} prodList={prodList} calcPrepCost={calcPrepCost} calcProductCost={calcProductCost}/>}
          {tab==="raw"       && (hasPerm("raw","view") ? <RawTab t={t} lang={lang} C={DARK} rawList={rawList} setRawList={setRawList} classes={classes} prepList={prepList} prodList={prodList} showToast={showToast} hasPerm={hasPerm} mod="raw"/> : <NoPerm t={t} C={DARK}/>)}
          {tab==="prep"      && (hasPerm("prep","view") ? <PrepTab t={t} lang={lang} C={DARK} prepList={prepList} setPrepList={setPrepList} rawList={rawList} prodList={prodList} classes={classes} calcPrepCost={calcPrepCost} showToast={showToast} hasPerm={hasPerm} mod="prep"/> : <NoPerm t={t} C={DARK}/>)}
          {tab==="products"  && (hasPerm("products","view") ? <ProductsTab t={t} lang={lang} C={DARK} prodList={prodList} setProdList={setProdList} rawList={rawList} prepList={prepList} classes={classes} calcPrepCost={calcPrepCost} calcProductCost={calcProductCost} showToast={showToast} hasPerm={hasPerm} mod="products"/> : <NoPerm t={t} C={DARK}/>)}
          {tab==="classes"   && (hasPerm("classes","view") ? <ClassesTab t={t} lang={lang} C={DARK} classes={classes} setClasses={setClasses} showToast={showToast} hasPerm={hasPerm} mod="classes"/> : <NoPerm t={t} C={DARK}/>)}
          {tab==="users"     && currentUser.role==="admin" && <UsersTab t={t} lang={lang} C={DARK} users={users} setUsers={setUsers} showToast={showToast} currentUserId={currentUser.id}/>}
        </div>
      </div>

      {toast && <div className={`toast ${toast.type}`}>{toast.msg}</div>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// LOGIN
// ═══════════════════════════════════════════════════════════════
function LoginScreen({users,onLogin}) {
  const [uid,setUid]     = useState("");
  const [pin,setPin]     = useState("");
  const [err,setErr]     = useState("");
  const [shake,setShake] = useState(false);
  const [loading,setLoading] = useState(false);
  const [focused,setFocused] = useState(null);
  const [lang,setLang]   = useState("ar");
  const t = T[lang];
  const isRtl = lang === "ar";

  const handleLogin = () => {
    const u = users.find(x=>x.id===uid.trim()&&x.pin===pin.trim());
    if(u) {
      setLoading(true);
      setTimeout(()=>onLogin(u), 800);
    } else {
      setErr(t.loginErr);
      setShake(true);
      setTimeout(()=>setShake(false), 600);
    }
  };

  // PIN dots display
  const pinDots = Array.from({length:5},(_,i)=>i < pin.length);

  return (
    <div dir={isRtl?"rtl":"ltr"} style={{
      minHeight:"100vh", fontFamily: isRtl ? "'Cairo',sans-serif" : "'DM Sans',sans-serif",
      background:"#060810", display:"flex", overflow:"hidden", position:"relative"
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700;900&family=DM+Sans:wght@300;400;500;600;700&family=Syne:wght@700;800&display=swap');
        @keyframes fadeUp{from{opacity:0;transform:translateY(28px)}to{opacity:1;transform:translateY(0)}}
        @keyframes shakeX{0%,100%{transform:translateX(0)}15%{transform:translateX(-10px)}30%{transform:translateX(10px)}45%{transform:translateX(-8px)}60%{transform:translateX(8px)}75%{transform:translateX(-4px)}90%{transform:translateX(4px)}}
        @keyframes spinIn{from{opacity:0;transform:scale(.6) rotate(-20deg)}to{opacity:1;transform:scale(1) rotate(0)}}
        @keyframes gridFloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-18px)}}
        @keyframes pulse{0%,100%{opacity:.4;transform:scale(1)}50%{opacity:.7;transform:scale(1.04)}}
        @keyframes lineGrow{from{width:0}to{width:100%}}
        @keyframes dotPop{from{transform:scale(0)}to{transform:scale(1)}}
        @keyframes loadBar{from{width:0}to{width:100%}}
        .lf-wrap{animation:fadeUp .55s ease both}
        .lf-card{animation:fadeUp .55s .1s ease both}
        .lf-shake{animation:shakeX .55s ease}
        .lf-input-wrap{position:relative;margin-bottom:22px}
        .lf-input{
          width:100%;background:transparent;border:none;border-bottom:2px solid #1e2540;
          color:#dde3f0;padding:14px 0 10px;font-size:15px;outline:none;
          font-family:inherit;transition:border-color .2s;
          letter-spacing:.5px;
        }
        .lf-input:focus{border-bottom-color:#e8a020}
        .lf-input::placeholder{color:#2a3350;font-size:14px}
        .lf-label{
          position:absolute;top:14px;${isRtl?"right":"left"}:0;
          font-size:12px;font-weight:700;color:#3a4a6a;
          text-transform:uppercase;letter-spacing:.1em;
          transition:all .18s;pointer-events:none;
        }
        .lf-input:focus ~ .lf-label,
        .lf-input:not(:placeholder-shown) ~ .lf-label{
          top:-4px;font-size:10px;color:#e8a020;letter-spacing:.12em;
        }
        .lf-line{
          position:absolute;bottom:0;${isRtl?"right":"left"}:0;
          height:2px;background:#e8a020;width:0;
          transition:width .25s ease;
        }
        .lf-input:focus ~ .lf-label ~ .lf-line{width:100%}
        .lf-btn{
          width:100%;background:linear-gradient(135deg,#e8a020,#c4861a);
          color:#060810;border:none;border-radius:10px;
          padding:14px;font-size:14px;font-weight:800;
          font-family:inherit;cursor:pointer;letter-spacing:.08em;
          text-transform:uppercase;transition:all .2s;
          position:relative;overflow:hidden;
        }
        .lf-btn:hover{filter:brightness(1.1);box-shadow:0 8px 32px #e8a02044;transform:translateY(-1px)}
        .lf-btn:active{transform:translateY(0);filter:brightness(.98)}
        .lf-btn-loading{pointer-events:none;opacity:.8}
        .lf-loadbar{position:absolute;bottom:0;left:0;height:3px;background:rgba(0,0,0,.3);animation:loadBar .8s ease forwards}
        .lf-err{color:#f87171;font-size:12px;text-align:center;margin-bottom:16px;font-weight:600;animation:fadeUp .2s ease}
        .lf-lang{background:transparent;border:1px solid #1e2540;color:#5a6585;padding:6px 18px;border-radius:20px;cursor:pointer;font-size:12px;font-weight:700;font-family:inherit;letter-spacing:.06em;transition:all .18s}
        .lf-lang:hover{border-color:#e8a02066;color:#e8a020}
        .dot{width:12px;height:12px;border-radius:50%;border:2px solid #3a4a6a;transition:all .2s;display:inline-block;margin:0 4px}
        .dot.filled{background:#e8a020;border-color:#e8a020;animation:dotPop .15s ease;box-shadow:0 0 10px #e8a02066}
      `}</style>

      {/* ── BACKGROUND GRID ── */}
      <div style={{position:"absolute",inset:0,overflow:"hidden",pointerEvents:"none"}}>
        {/* deep gradient */}
        <div style={{position:"absolute",inset:0,background:"radial-gradient(ellipse 80% 60% at 50% 0%,#0d1428 0%,#060810 70%)"}}/>
        {/* grid lines */}
        <svg width="100%" height="100%" style={{position:"absolute",inset:0,opacity:.07}}>
          <defs>
            <pattern id="grid" width="60" height="60" patternUnits="userSpaceOnUse">
              <path d="M 60 0 L 0 0 0 60" fill="none" stroke="#e8a020" strokeWidth=".8"/>
            </pattern>
          </defs>
          <rect width="100%" height="100%" fill="url(#grid)"/>
        </svg>
        {/* glow orbs */}
        <div style={{position:"absolute",top:"-15%",left:"60%",width:500,height:500,background:"radial-gradient(circle,#e8a02018 0%,transparent 70%)",animation:"pulse 6s ease-in-out infinite"}}/>
        <div style={{position:"absolute",bottom:"-10%",left:"-10%",width:400,height:400,background:"radial-gradient(circle,#3b82f614 0%,transparent 70%)",animation:"pulse 8s ease-in-out infinite 2s"}}/>

      </div>

      {/* ── SPLIT LAYOUT ── */}
      <div style={{display:"flex",width:"100%",position:"relative",zIndex:1}}>

        {/* LEFT BRAND PANEL */}
        <div style={{
          flex:1, flexDirection:"column", justifyContent:"center",
          padding:"60px 80px", display: "flex"
        }}>
          <div className="lf-wrap">
            <div style={{
              fontFamily:"'Syne',sans-serif", fontSize:52, fontWeight:800,
              color:"#dde3f0", lineHeight:1.05, marginBottom:16,
              textShadow:"0 0 80px #e8a02022"
            }}>
              TALA<br/><span style={{color:"#e8a020"}}>COSTING</span>
            </div>
            <div style={{width:48,height:3,background:"linear-gradient(90deg,#e8a020,transparent)",marginBottom:20,borderRadius:2}}/>
            <div style={{fontSize:14,color:"#5a6585",lineHeight:1.8,maxWidth:320,fontWeight:400}}>
              {isRtl
                ? "نظام متكامل لإدارة وتحليل تكاليف المنتجات والمواد"
                : "Integrated system for managing and analyzing product & material costs"}
            </div>

          </div>
        </div>

        {/* RIGHT LOGIN PANEL */}
        <div style={{
          width: 480,
          minWidth: 480,
          display:"flex", flexDirection:"column", justifyContent:"center",
          padding: "60px 56px",
          background:"#0a0d1a",
          borderLeft: "1px solid #1e2540",
          position:"relative"
        }}>
          {/* corner accent */}
          <div style={{position:"absolute",top:0,left:0,width:80,height:3,background:"linear-gradient(90deg,#e8a020,transparent)"}}/>
          <div style={{position:"absolute",top:0,left:0,width:3,height:80,background:"linear-gradient(180deg,#e8a020,transparent)"}}/>

          <div className="lf-card">
            {/* Mobile logo */}
            <div style={{marginBottom:36}}>
              <div style={{fontFamily:"'Syne',sans-serif",fontSize:26,fontWeight:800,color:"#dde3f0",letterSpacing:"1px"}}>
                TALA <span style={{color:"#e8a020"}}>COSTING</span>
              </div>
              <div style={{fontSize:12,color:"#3a4a6a",marginTop:4,fontWeight:600,letterSpacing:".1em",textTransform:"uppercase"}}>
                {isRtl ? "تسجيل الدخول" : "Sign in to continue"}
              </div>
            </div>

            {/* User ID field */}
            <div className={`lf-input-wrap${shake?" lf-shake":""}`}>
              <input
                className="lf-input" value={uid} placeholder=" "
                onChange={e=>setUid(e.target.value)}
                onFocus={()=>setFocused("uid")} onBlur={()=>setFocused(null)}
                onKeyDown={e=>e.key==="Enter"&&handleLogin()}
                style={{textAlign:isRtl?"right":"left"}}
              />
              <span className="lf-label">{t.userId}</span>
              <div className="lf-line"/>
            </div>

            {/* PIN field — dots */}
            <div className={`lf-input-wrap${shake?" lf-shake":""}`} style={{marginBottom:10}}>
              <input
                className="lf-input" type="password" maxLength={5} value={pin} placeholder=" "
                onChange={e=>setPin(e.target.value.replace(/\D/g,""))}
                onFocus={()=>setFocused("pin")} onBlur={()=>setFocused(null)}
                onKeyDown={e=>e.key==="Enter"&&handleLogin()}
                style={{opacity:0,position:"absolute",width:"100%",cursor:"default"}}
              />
              <span className="lf-label" style={{top:focused==="pin"||pin.length>0?"-4px":"14px",fontSize:focused==="pin"||pin.length>0?"10px":"12px",color:focused==="pin"?"#e8a020":"#3a4a6a"}}>{t.pin}</span>
              {/* visual dots */}
              <div
                style={{
                  borderBottom:`2px solid ${focused==="pin"?"#e8a020":"#1e2540"}`,
                  padding:"14px 0 14px",display:"flex",
                  gap:10, alignItems:"center",
                  justifyContent:isRtl?"flex-end":"flex-start",
                  cursor:"text", transition:"border-color .2s"
                }}
                onClick={()=>{ document.querySelector('input[type="password"]').focus(); }}
              >
                {pinDots.map((filled,i)=>(
                  <span key={i} className={`dot${filled?" filled":""}`}/>
                ))}
              </div>
            </div>

            {err && <div className="lf-err">{err}</div>}

            <button
              className={`lf-btn${loading?" lf-btn-loading":""}`}
              onClick={handleLogin} style={{marginTop:12}}
            >
              {loading ? (
                <span style={{display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
                  <span style={{width:16,height:16,border:"2px solid #06081088",borderTop:"2px solid #060810",borderRadius:"50%",display:"inline-block",animation:"spin .7s linear infinite"}}/>
                  {isRtl?"جاري الدخول...":"Signing in..."}
                  <div className="lf-loadbar"/>
                </span>
              ) : t.loginBtn}
            </button>

            <div style={{display:"flex",justifyContent:"center",marginTop:24}}>
              <button className="lf-lang" onClick={()=>setLang(l=>l==="ar"?"en":"ar")}>{t.lang}</button>
            </div>
          </div>
        </div>
      </div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

function NoPerm({t,C=DARK}) {
  return <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:200,color:C.muted,fontSize:14}}>{t.noPermission}</div>;
}

// ═══════════════════════════════════════════════════════════════
// DASHBOARD / ANALYTICS
// ═══════════════════════════════════════════════════════════════
function DashboardTab({t,lang,C,rawList,prepList,prodList,calcPrepCost,calcProductCost}) {
  const [topN,setTopN] = useState(10);

  const prodCalc = prodList.map(p=>({...p,...calcProductCost(p)}));
  const prepCalc = prepList.map(p=>({...p,...calcPrepCost(p)}));

  // KPIs
  const totalProds = prodList.length;
  const totalPreps = prepList.length;
  const totalRaws  = rawList.length;
  const avgMargin  = prodCalc.filter(p=>p.sellingPrice>0).length
    ? prodCalc.filter(p=>p.sellingPrice>0).reduce((a,p)=>a+p.margin,0)/prodCalc.filter(p=>p.sellingPrice>0).length : 0;
  const avgCost    = prodCalc.length ? prodCalc.reduce((a,p)=>a+p.totalCost,0)/prodCalc.length : 0;

  // Variance analysis
  const withStd    = prodCalc.filter(p=>p.stdCost>0);
  const overBudget = withStd.filter(p=>p.totalCost>p.stdCost);
  const underBudget= withStd.filter(p=>p.totalCost<=p.stdCost);

  // Margin distribution
  const highMargin = prodCalc.filter(p=>p.margin>30&&p.sellingPrice>0).length;
  const midMargin  = prodCalc.filter(p=>p.margin>=15&&p.margin<=30&&p.sellingPrice>0).length;
  const lowMargin  = prodCalc.filter(p=>p.margin<15&&p.sellingPrice>0&&p.margin>0).length;
  const noPrice    = prodCalc.filter(p=>!p.sellingPrice||p.sellingPrice===0).length;

  // Top cost drivers across all products
  const driverMap = {};
  prodList.forEach(prod=>{
    prod.ingredients?.forEach(ing=>{
      if(ing.source==="raw"){
        const raw=rawList.find(r=>String(r.id)===String(ing.srcId));
        if(!raw) return;
        const qty=parseFloat(ing.qty)||0;
        const cost=(qty/1000)*raw.price;
        driverMap[raw.name]=(driverMap[raw.name]||0)+cost;
      } else {
        const prep=prepList.find(p=>String(p.id)===String(ing.srcId));
        if(!prep) return;
        const {costPerUnit}=calcPrepCost(prep);
        const qty=parseFloat(ing.qty)||0;
        const cost=(qty/1000)*costPerUnit;
        driverMap[prep.name]=(driverMap[prep.name]||0)+cost;
      }
    });
  });
  const drivers = Object.entries(driverMap).sort((a,b)=>b[1]-a[1]).slice(0,topN);
  const maxDriver = drivers[0]?.[1]||1;

  // Prep usage
  const prepUsage = prepCalc.map(p=>({
    ...p, usedIn:prodList.filter(pr=>pr.ingredients?.some(i=>i.source==="prep"&&String(i.srcId)===String(p.id))).length
  }));

  // Raw usage
  const rawUsage = rawList.map(r=>({
    ...r,
    inPrep:prepList.filter(p=>p.ingredients?.some(i=>String(i.rawId)===String(r.id))).length,
    inProd:prodList.filter(p=>p.ingredients?.some(i=>i.source==="raw"&&String(i.srcId)===String(r.id))).length,
  }));

  const Card = ({label,value,sub,color})=>(
    <div className="card" style={{padding:"14px 18px"}}>
      <div style={{fontSize:24,fontWeight:800,color:color||DARK.accent}}>{value}</div>
      <div style={{fontSize:12,color:DARK.muted,marginTop:3}}>{label}</div>
      {sub&&<div style={{fontSize:11,color:DARK.muted,marginTop:2}}>{sub}</div>}
    </div>
  );

  const SectionTitle = ({children})=>(
    <div style={{fontWeight:700,fontSize:14,color:DARK.text,marginBottom:12,paddingBottom:8,borderBottom:"1px solid "+DARK.border}}>{children}</div>
  );

  const Bar = ({val,max,color})=>(
    <div style={{background:DARK.surface,borderRadius:3,height:6,flex:1}}>
      <div style={{width:(val/max*100)+"%",height:6,borderRadius:3,background:color,transition:"width .4s ease"}}/>
    </div>
  );

  const noDataMsg = <div style={{color:DARK.muted,fontSize:13,textAlign:"center",padding:"24px 0"}}>{t.noData}</div>;

  return (
    <div style={{maxWidth:1100,margin:"0 auto"}}>
      {/* Top N selector */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:18,justifyContent:"flex-end"}}>
        <span style={{fontSize:12,color:DARK.muted}}>{t.show}</span>
        {[5,10,20].map(n=>(
          <button key={n} className={`top-n-btn${topN===n?" active":""}`} onClick={()=>setTopN(n)}>{n}</button>
        ))}
      </div>

      {/* KPI CARDS */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,marginBottom:20}}>
        <Card label={t.totalProducts}  value={totalProds}               color={DARK.accent}/>
        <Card label={t.totalPrep}      value={totalPreps}               color={DARK.blue}/>
        <Card label={t.totalRaw}       value={totalRaws}                color="#a78bfa"/>
        <Card label={t.avgMarginLbl}   value={avgMargin.toFixed(1)+"%"} color={avgMargin>30?DARK.green:avgMargin>15?DARK.yellow:DARK.red}/>
        <Card label={t.avgCostLbl}     value={avgCost.toFixed(2)}       color="#f87171"/>
        <Card label={t.productsOverBudget} value={overBudget.length} sub={withStd.length>0?`/ ${withStd.length} ${lang==="ar"?"لهم معياري":"with target"}`:""} color={overBudget.length>0?DARK.red:DARK.green}/>
      </div>

      {/* MARGIN DISTRIBUTION */}
      <div className="card" style={{padding:18,marginBottom:16}}>
        <SectionTitle>{t.marginDistribution}</SectionTitle>
        {prodCalc.length===0?noDataMsg:(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10}}>
            {[
              {label:t.highMargin, count:highMargin, color:DARK.green,   pct:totalProds?((highMargin/totalProds)*100).toFixed(0):0},
              {label:t.midMargin,  count:midMargin,  color:DARK.yellow,  pct:totalProds?((midMargin/totalProds)*100).toFixed(0):0},
              {label:t.lowMargin,  count:lowMargin,  color:DARK.red,     pct:totalProds?((lowMargin/totalProds)*100).toFixed(0):0},
              {label:t.noSelling,  count:noPrice,    color:DARK.muted,   pct:totalProds?((noPrice/totalProds)*100).toFixed(0):0},
            ].map((s,i)=>(
              <div key={i} style={{background:DARK.surface,borderRadius:10,padding:"12px 14px",border:"1px solid "+DARK.border}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                  <span style={{fontSize:11,color:DARK.muted,fontWeight:600}}>{s.label}</span>
                  <span style={{fontSize:11,color:s.color,fontWeight:700}}>{s.pct}%</span>
                </div>
                <div style={{fontSize:22,fontWeight:800,color:s.color}}>{s.count}</div>
                <div style={{background:DARK.border,borderRadius:3,height:4,marginTop:8}}>
                  <div style={{width:s.pct+"%",height:4,borderRadius:3,background:s.color}}/>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* VARIANCE REPORT */}
      {withStd.length>0&&(
        <div className="card" style={{padding:18,marginBottom:16}}>
          <SectionTitle>{t.varianceReport}</SectionTitle>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>
                {[t.productName,t.totalCost,t.stdCost,t.variance,t.variancePct,t.topCostDriver].map((h,i)=>(
                  <th key={i} style={{padding:"9px 12px",textAlign:lang==="ar"?"right":"left",fontSize:10,fontWeight:700,color:DARK.muted,textTransform:"uppercase",letterSpacing:".06em",background:DARK.surface,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {[...withStd].sort((a,b)=>(b.totalCost-b.stdCost)-(a.totalCost-a.stdCost)).slice(0,topN).map((p,i)=>{
                  const v=p.totalCost-p.stdCost;
                  const vp=((p.totalCost-p.stdCost)/p.stdCost)*100;
                  // find top cost driver
                  let topDriver="—"; let topCost=0;
                  p.ingredients?.forEach(ing=>{
                    let cost=0,name="";
                    if(ing.source==="raw"){const r=rawList.find(r=>String(r.id)===String(ing.srcId));if(r){cost=(parseFloat(ing.qty)||0)/1000*r.price;name=r.name;}}
                    else{const pr=prepList.find(pr=>String(pr.id)===String(ing.srcId));if(pr){const {costPerUnit}=calcPrepCost(pr);cost=(parseFloat(ing.qty)||0)/1000*costPerUnit;name=pr.name;}}
                    if(cost>topCost){topCost=cost;topDriver=name;}
                  });
                  return(
                    <tr key={p.id} style={{borderBottom:"1px solid "+DARK.border+"22",background:v>0?DARK.red+"08":"transparent"}}>
                      <td style={{padding:"10px 12px",fontWeight:600}}>{p.name}</td>
                      <td style={{padding:"10px 12px",color:"#f87171",fontWeight:600}}>{p.totalCost.toFixed(2)}</td>
                      <td style={{padding:"10px 12px",color:DARK.muted}}>{p.stdCost.toFixed(2)}</td>
                      <td style={{padding:"10px 12px"}}><span style={{color:v>0?"#f87171":"#4ade80",fontWeight:700}}>{v>0?"+":""}{v.toFixed(2)}</span></td>
                      <td style={{padding:"10px 12px"}}><span style={{background:v>0?"#f8717122":"#4ade8022",color:v>0?"#f87171":"#4ade80",padding:"2px 8px",borderRadius:20,fontSize:12,fontWeight:700}}>{vp>0?"+":""}{vp.toFixed(1)}%</span></td>
                      <td style={{padding:"10px 12px",color:DARK.yellow,fontSize:12,fontWeight:600}}>{topDriver}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* PRODUCTS ANALYSIS */}
      <div className="card" style={{padding:18,marginBottom:16}}>
        <SectionTitle>{t.secProducts}</SectionTitle>
        {prodCalc.length===0?noDataMsg:(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:16}}>
            {[
              {title:lang==="ar"?"أعلى هامش ربح":"Highest Margin", data:[...prodCalc].filter(p=>p.sellingPrice>0).sort((a,b)=>b.margin-a.margin), valFn:p=>p.margin, fmt:v=>v.toFixed(1)+"%", color:p=>p.margin>30?DARK.green:p.margin>15?DARK.yellow:DARK.red},
              {title:lang==="ar"?"أعلى تكلفة":"Highest Cost",    data:[...prodCalc].sort((a,b)=>b.totalCost-a.totalCost),           valFn:p=>p.totalCost, fmt:v=>v.toFixed(2), color:()=>"#f87171"},
              {title:lang==="ar"?"أقل تكلفة":"Lowest Cost",      data:[...prodCalc].filter(p=>p.totalCost>0).sort((a,b)=>a.totalCost-b.totalCost), valFn:p=>p.totalCost, fmt:v=>v.toFixed(2), color:()=>DARK.green},
            ].map((sec,si)=>(
              <div key={si}>
                <div style={{fontSize:11,color:DARK.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",marginBottom:10}}>{sec.title}</div>
                {sec.data.slice(0,topN).map((p,i)=>{
                  const val=sec.valFn(p);
                  const max=sec.valFn(sec.data[0])||1;
                  const col=sec.color(p);
                  return(
                    <div key={p.id} style={{marginBottom:8}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                        <span style={{fontSize:12,color:DARK.text,fontWeight:600,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{i+1}. {p.name}</span>
                        <span style={{fontSize:12,color:col,fontWeight:700,marginRight:6,minWidth:60,textAlign:"left"}}>{sec.fmt(val)}</span>
                      </div>
                      <Bar val={val} max={max} color={col}/>
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* TOP COST DRIVERS */}
      <div className="card" style={{padding:18,marginBottom:16}}>
        <SectionTitle>{t.topCostDriversReport}</SectionTitle>
        {drivers.length===0?noDataMsg:(
          <div>
            {drivers.map(([name,cost],i)=>(
              <div key={i} style={{marginBottom:9}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                  <span style={{fontSize:12,fontWeight:600,color:DARK.text}}>{i+1}. {name}</span>
                  <span style={{fontSize:12,color:DARK.yellow,fontWeight:700}}>{cost.toFixed(2)}</span>
                </div>
                <Bar val={cost} max={maxDriver} color={DARK.yellow}/>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* PREP ANALYSIS */}
      <div className="card" style={{padding:18,marginBottom:16}}>
        <SectionTitle>{t.secPrep}</SectionTitle>
        {prepCalc.length===0?noDataMsg:(
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
            <div>
              <div style={{fontSize:11,color:DARK.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",marginBottom:10}}>{t.prepHighCost}</div>
              {[...prepCalc].sort((a,b)=>b.costPerUnit-a.costPerUnit).slice(0,topN).map((p,i)=>(
                <div key={p.id} style={{marginBottom:8}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                    <span style={{fontSize:12,color:DARK.text,fontWeight:600}}>{i+1}. {p.name}</span>
                    <span style={{fontSize:12,color:DARK.accent,fontWeight:700}}>{p.costPerUnit.toFixed(3)}</span>
                  </div>
                  <Bar val={p.costPerUnit} max={prepCalc[0]?.costPerUnit||1} color={DARK.accent}/>
                </div>
              ))}
            </div>
            <div>
              <div style={{fontSize:11,color:DARK.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",marginBottom:10}}>{t.prepMostUsed}</div>
              {[...prepUsage].sort((a,b)=>b.usedIn-a.usedIn).slice(0,topN).map((p,i)=>(
                <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:DARK.surface,borderRadius:8,marginBottom:6,border:"1px solid "+DARK.border}}>
                  <span style={{fontSize:12,fontWeight:600,color:DARK.text}}>{i+1}. {p.name}</span>
                  <span style={{background:DARK.blue+"22",color:DARK.blue,padding:"2px 10px",borderRadius:20,fontSize:12,fontWeight:700}}>{p.usedIn} {t.productsCount}</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* RAW MATERIALS ANALYSIS */}
      <div className="card" style={{padding:18,marginBottom:16}}>
        <SectionTitle>{t.secRaw}</SectionTitle>
        {rawList.length===0?noDataMsg:(
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
            <div>
              <div style={{fontSize:11,color:DARK.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",marginBottom:10}}>{t.rawHighPrice}</div>
              {[...rawUsage].sort((a,b)=>b.price-a.price).slice(0,topN).map((r,i)=>(
                <div key={r.id} style={{marginBottom:8}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                    <span style={{fontSize:12,color:DARK.text,fontWeight:600}}>{i+1}. {r.name}</span>
                    <span style={{fontSize:12,color:DARK.accent,fontWeight:700}}>{r.price.toFixed(2)}</span>
                  </div>
                  <Bar val={r.price} max={rawUsage[0]?.price||1} color={DARK.accent}/>
                </div>
              ))}
            </div>
            <div>
              <div style={{fontSize:11,color:DARK.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",marginBottom:10}}>{t.rawMostUsed}</div>
              {[...rawUsage].sort((a,b)=>(b.inPrep+b.inProd)-(a.inPrep+a.inProd)).slice(0,topN).map((r,i)=>(
                <div key={r.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:DARK.surface,borderRadius:8,marginBottom:6,border:"1px solid "+DARK.border}}>
                  <span style={{fontSize:12,fontWeight:600,color:DARK.text}}>{i+1}. {r.name}</span>
                  <div style={{display:"flex",gap:6}}>
                    <span style={{background:"#1e3a5f",color:"#60a5fa",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700}}>P:{r.inPrep}</span>
                    <span style={{background:"#0a3326",color:"#4ade80",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700}}>M:{r.inProd}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// RAW MATERIALS
// ═══════════════════════════════════════════════════════════════
function RawTab({t,lang,C=DARK,rawList,setRawList,classes,prepList=[],prodList=[],showToast,hasPerm,mod}) {
  const [search,setSearch]=useState(""); const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",unit:"kg",price:"",class:""}); const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null); const [usageModal,setUsageModal]=useState(null); const fileRef=useRef();
  const cls=classes.raw||[];
  const ok=()=>{ const e={}; if(!form.name.trim())e.name=t.required; if(!form.price)e.price=t.required; else if(parseFloat(form.price)<=0)e.price=t.positiveNum; setErrs(e); return !Object.keys(e).length; };
  const reset=()=>{ setForm({name:"",unit:"kg",price:"",class:""}); setErrs({}); setShowForm(false); setEditId(null); };
  const save=()=>{ if(!ok()) return; const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US"); if(editId!==null) setRawList(p=>p.map(m=>m.id===editId?{...m,name:form.name.trim(),unit:form.unit,price:parseFloat(form.price),class:form.class,lastUpdated:now}:m)); else { const code=genCode("Raw",rawList); setRawList(p=>[...p,{id:Date.now(),code,name:form.name.trim(),unit:form.unit,price:parseFloat(form.price),class:form.class,lastUpdated:now}]); } reset(); showToast(t.savedOk); };
  const doEdit=m=>{ setForm({name:m.name,unit:m.unit,price:String(m.price),class:m.class||""}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setRawList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };
  const doExport=()=>{ const ws=XLSX.utils.json_to_sheet(rawList.map((m,i)=>({"#":i+1,[t.code]:m.code,[t.name]:m.name,[t.class]:m.class||"",[t.unit]:m.unit,[t.price]:m.price}))); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,t.rawMat); XLSX.writeFile(wb,`raw_${Date.now()}.xlsx`); };
  const doImport=e=>{ const file=e.target.files[0]; if(!file) return; const r=new FileReader(); r.onload=ev=>{ try{ const wb=XLSX.read(ev.target.result,{type:"binary"}); const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]); let n=0; const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US"); setRawList(prev=>{ const u=[...prev]; rows.forEach(row=>{ const rn=String(row[t.name]||row.Name||row["الاسم"]||"").trim(); const rc=String(row[t.code]||row.Code||row["الكود"]||"").trim(); const rp=parseFloat(row[t.price]||row.Price||row["السعر"]||0); if(!rp||rp<=0) return; let i=u.findIndex(m=>m.name.toLowerCase()===rn.toLowerCase()); if(i===-1) i=u.findIndex(m=>m.code===rc); if(i!==-1){u[i]={...u[i],price:rp,lastUpdated:now};n++;} }); return u; }); showToast(n>0?`${t.importedOk} ${n} ${t.importedItems}`:t.noMatch,n>0?"success":"warning"); }catch{ showToast(lang==="ar"?"خطأ في الملف":"File error","error"); } }; r.readAsBinaryString(file); e.target.value=""; };
  const filtered=rawList.filter(m=>(m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase()))&&(fcls==="all"||m.class===fcls));
  return (
    <div>
      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:220}} placeholder={t.search} value={search} onChange={e=>setSearch(e.target.value)}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
          {hasPerm(mod,"edit")&&<><button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button><button className="btn btn-secondary" onClick={()=>fileRef.current.click()}>{t.importXlsx}</button><input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={doImport}/></>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>
      <div style={{display:"flex",gap:7,marginBottom:10,flexWrap:"wrap"}}>
        <button className={`filter-btn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all}</button>
        {cls.map(c=><button key={c} className={`filter-btn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c}</button>)}
      </div>
      <div style={{fontSize:11,color:"#2e3a55",marginBottom:10}}>* {t.importNote}</div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>{["#",t.code,t.name,t.class,t.unit,t.price,hasPerm(mod,"edit")||hasPerm(mod,"delete")?t.actions:""].filter(Boolean).map((h,i)=><th key={i}>{h}</th>)}</tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={7} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :filtered.map((m,i)=>(
              <tr key={m.id}>
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>{m.name}</td>
                <td><span className="badge badge-cls">{m.class||"—"}</span></td>
                <td><span className={`badge badge-${m.unit}`}>{unitLbl(m.unit,t)}</span></td>
                <td style={{color:C.accent,fontWeight:700}}>{m.price.toFixed(2)}</td>
                <td>{(()=>{const n=prepList.filter(p=>p.ingredients?.some(i=>String(i.rawId)===String(m.id))).length;return n>0?<button onClick={()=>setUsageModal({item:m,type:"prep",list:prepList.filter(p=>p.ingredients?.some(i=>String(i.rawId)===String(m.id)))})} style={{background:"#0f2a4a",color:"#60a5fa",border:"1px solid #1e3a6033",borderRadius:6,padding:"3px 10px",cursor:"pointer",fontWeight:700,fontSize:13}}>{n}</button>:<span style={{color:DARK.muted}}>0</span>;})()}</td>
                <td>{(()=>{const n=prodList.filter(p=>p.ingredients?.some(i=>i.source==="raw"&&String(i.srcId)===String(m.id))).length;return n>0?<button onClick={()=>setUsageModal({item:m,type:"product",list:prodList.filter(p=>p.ingredients?.some(i=>i.source==="raw"&&String(i.srcId)===String(m.id)))})} style={{background:"#0a2a1a",color:"#4ade80",border:"1px solid #16a34a33",borderRadius:6,padding:"3px 10px",cursor:"pointer",fontWeight:700,fontSize:13}}>{n}</button>:<span style={{color:DARK.muted}}>0</span>;})()}</td>
                {(hasPerm(mod,"edit")||hasPerm(mod,"delete"))&&<td><div style={{display:"flex",gap:5}}>{hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>doEdit(m)}>{t.edit}</button>}{hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>setDelId(m.id)}>{t.delete}</button>}</div></td>}
              </tr>
            ))}
          </tbody>
        </table>
      </div></div>

      {/* USAGE MODAL */}
      {usageModal&&(
        <div className="overlay" onClick={e=>e.target===e.currentTarget&&setUsageModal(null)}>
          <div className="modal" style={{maxWidth:500}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div>
                <div style={{fontWeight:800,fontSize:15,color:DARK.accent}}>{usageModal.item.name}</div>
                <div style={{fontSize:12,color:DARK.muted,marginTop:2}}>{usageModal.type==="prep"?t.relatedPreps:t.relatedProducts} ({usageModal.list.length})</div>
              </div>
              <button onClick={()=>setUsageModal(null)} style={{background:"transparent",border:"none",color:DARK.muted,fontSize:20,cursor:"pointer"}}>✕</button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {usageModal.list.map((item,i)=>(
                <div key={item.id} style={{background:DARK.surface,borderRadius:9,padding:"10px 14px",border:"1px solid "+DARK.border,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <code style={{background:DARK.bg,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8",marginLeft:8}}>{item.code}</code>
                    <span style={{fontWeight:600,fontSize:13}}>{item.name}</span>
                  </div>
                  <span className="badge badge-cls">{item.class||"—"}</span>
                </div>
              ))}
            </div>
            <div style={{display:"flex",justifyContent:"flex-end",marginTop:14}}>
              <button className="btn btn-secondary" onClick={()=>setUsageModal(null)}>{t.cancel}</button>
            </div>
          </div>
        </div>
      )}

      {showForm&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&reset()}><div className="modal">
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.edit:t.add} — {t.rawMat}</h2>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <div><label className="lbl">{t.name}</label><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">{t.class}</label><select value={form.class} onChange={e=>setForm({...form,class:e.target.value})}><option value="">—</option>{cls.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><label className="lbl">{t.unit}</label><select value={form.unit} onChange={e=>setForm({...form,unit:e.target.value})}><option value="kg">{t.kg}</option><option value="liter">{t.liter}</option><option value="piece">{t.piece}</option></select></div>
            <div><label className="lbl">{t.price}</label><input type="number" min="0" step="0.01" value={form.price} onChange={e=>setForm({...form,price:e.target.value})} placeholder="0.00"/>{errs.price&&<div className="err">{errs.price}</div>}</div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:4}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
        </div>
      </div></div>}
      {delId&&<DelModal t={t} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// PREP ITEMS
// ═══════════════════════════════════════════════════════════════
function PrepTab({t,lang,C=DARK,prepList,setPrepList,rawList,prodList=[],classes,calcPrepCost,showToast,hasPerm,mod}) {
  const [search,setSearch]=useState(""); const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",unit:"kg",class:"",yieldOverride:"",ingredients:[]}); const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null); const fileRef=useRef();
  const [ingSearch,setIngSearch]=useState("");
  const [viewItem,setViewItem]=useState(null);
  const [usageModal,setUsageModal]=useState(null);
  const cls=classes.prep||[];
  const blank=()=>({id:Date.now()+Math.random(),source:"raw",rawId:"",qty:"",waste:"0"});
  const reset=()=>{ setForm({name:"",unit:"kg",class:"",yieldOverride:"",ingredients:[]}); setErrs({}); setShowForm(false); setEditId(null); };
  const ok=()=>{ const e={}; if(!form.name.trim())e.name=t.required; setErrs(e); return !Object.keys(e).length; };
  const save=()=>{ if(!ok()) return; const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US"); const ings=form.ingredients.filter(i=>i.rawId&&parseFloat(i.qty)>0); if(editId!==null) setPrepList(p=>p.map(m=>m.id===editId?{...m,name:form.name.trim(),unit:form.unit,class:form.class,yieldOverride:form.yieldOverride,ingredients:ings,lastUpdated:now}:m)); else { const code=genCode("Prep",prepList); setPrepList(p=>[...p,{id:Date.now(),code,name:form.name.trim(),unit:form.unit,class:form.class,yieldOverride:form.yieldOverride,ingredients:ings,lastUpdated:now}]); } reset(); showToast(t.savedOk); };
  const doEdit=m=>{ setForm({name:m.name,unit:m.unit,class:m.class||"",yieldOverride:m.yieldOverride||"",ingredients:m.ingredients||[]}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setPrepList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };
  const addI=()=>setForm(f=>({...f,ingredients:[...f.ingredients,blank()]}));
  const remI=id=>setForm(f=>({...f,ingredients:f.ingredients.filter(i=>i.id!==id)}));
  const updI=(id,k,v)=>setForm(f=>({...f,ingredients:f.ingredients.map(i=>i.id===id?{...i,[k]:v}:i)}));
  const live=calcPrepCost({ingredients:form.ingredients,unit:form.unit,yieldOverride:form.yieldOverride});
  const doExport=()=>{
    const rows=[];
    prepList.forEach(p=>{
      const ings=p.ingredients||[];
      if(ings.length===0){
        rows.push({
          [lang==="ar"?"كود Prep":"Prep Code"]:p.code,
          [lang==="ar"?"اسم Prep":"Prep Name"]:p.name,
          [lang==="ar"?"الكلاس":"Class"]:p.class||"",
          [lang==="ar"?"الوحدة":"Unit"]:p.unit,
          [lang==="ar"?"نوع المكون":"Ingredient Type"]:"",
          [lang==="ar"?"كود المكون":"Ingredient Code"]:"",
          [lang==="ar"?"اسم المكون":"Ingredient Name"]:"",
          [lang==="ar"?"الكمية (g/ml)":"Qty (g/ml)"]:"",
          [lang==="ar"?"الهدر %":"Waste %"]:"",
        });
      } else {
        ings.forEach(ing=>{
          let ingCode="", ingName="", ingType="";
          if(ing.source==="prep"){
            const p=prepList.find(p=>String(p.id)===String(ing.rawId));
            if(p){ ingCode=p.code; ingName=p.name; ingType=lang==="ar"?"بريب":"Prep"; }
          } else {
            const raw=rawList.find(r=>String(r.id)===String(ing.rawId));
            if(raw){ ingCode=raw.code; ingName=raw.name; ingType=lang==="ar"?"مادة خام":"Raw"; }
          }
          rows.push({
            [lang==="ar"?"كود Prep":"Prep Code"]:p.code,
            [lang==="ar"?"اسم Prep":"Prep Name"]:p.name,
            [lang==="ar"?"الكلاس":"Class"]:p.class||"",
            [lang==="ar"?"الوحدة":"Unit"]:p.unit,
            [lang==="ar"?"نوع المكون":"Ingredient Type"]:ingType,
            [lang==="ar"?"كود المكون":"Ingredient Code"]:ingCode,
            [lang==="ar"?"اسم المكون":"Ingredient Name"]:ingName,
            [lang==="ar"?"الكمية (g/ml)":"Qty (g/ml)"]:parseFloat(ing.qty)||0,
            [lang==="ar"?"الهدر %":"Waste %"]:parseFloat(ing.waste)||0,
          });
        });
      }
    });
    const ws=XLSX.utils.json_to_sheet(rows);
    // set column widths
    ws["!cols"]=[{wch:14},{wch:20},{wch:14},{wch:10},{wch:16},{wch:14},{wch:22},{wch:14},{wch:10}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,t.prepItem);
    XLSX.writeFile(wb,`prep_items_${Date.now()}.xlsx`);
  };
  const doImport=e=>{
    const file=e.target.files[0]; if(!file) return;
    const r=new FileReader();
    r.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        let updated=0;
        const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
        // group rows by prep code
        const prepGroups={};
        rows.forEach(row=>{
          const pc=String(row[lang==="ar"?"كود Prep":"Prep Code"]||row["Prep Code"]||row["كود Prep"]||"").trim();
          const pn=String(row[lang==="ar"?"اسم Prep":"Prep Name"]||row["Prep Name"]||row["اسم Prep"]||"").trim();
          const key=pc||pn;
          if(!key) return;
          if(!prepGroups[key]) prepGroups[key]={code:pc,name:pn,ingredients:[]};
          const ingCode=String(row[lang==="ar"?"كود المكون":"Ingredient Code"]||row["Ingredient Code"]||row["كود المكون"]||"").trim();
          const ingName=String(row[lang==="ar"?"اسم المكون":"Ingredient Name"]||row["Ingredient Name"]||row["اسم المكون"]||"").trim();
          const ingType=String(row[lang==="ar"?"نوع المكون":"Ingredient Type"]||row["Ingredient Type"]||row["نوع المكون"]||"raw").trim().toLowerCase();
          const qty=parseFloat(row[lang==="ar"?"الكمية (g/ml)":"Qty (g/ml)"]||row["Qty (g/ml)"]||row["الكمية (g/ml)"]||0);
          const waste=parseFloat(row[lang==="ar"?"الهدر %":"Waste %"]||row["Waste %"]||row["الهدر %"]||0);
          if(ingCode||ingName){
            // find raw material by code or name
            const isPrep = ingType.includes("prep") || ingType.includes("بريب");
            if(isPrep){
              const p=prepList.find(p=>p.code===ingCode||p.name.toLowerCase()===ingName.toLowerCase());
              if(p) prepGroups[key].ingredients.push({id:Date.now()+Math.random(),source:"prep",rawId:p.id,qty,waste});
            } else {
              const raw=rawList.find(r=>r.code===ingCode||r.name.toLowerCase()===ingName.toLowerCase());
              if(raw) prepGroups[key].ingredients.push({id:Date.now()+Math.random(),source:"raw",rawId:raw.id,qty,waste});
            }
          }
        });
        setPrepList(prev=>{
          const u=[...prev];
          Object.values(prepGroups).forEach(pg=>{
            // find prep by code then name
            let idx=u.findIndex(p=>p.code===pg.code);
            if(idx===-1) idx=u.findIndex(p=>p.name.toLowerCase()===pg.name.toLowerCase());
            if(idx===-1) return; // only update existing
            // smart sync ingredients
            const newIngs=pg.ingredients;
            u[idx]={...u[idx],ingredients:newIngs,lastUpdated:now};
            updated++;
          });
          return u;
        });
        showToast(updated>0?`${t.importedOk} ${updated} ${t.importedItems}`:t.noMatch,updated>0?"success":"warning");
      }catch(err){ showToast(lang==="ar"?"خطأ في الملف":"File error","error"); }
    };
    r.readAsBinaryString(file); e.target.value="";
  };
  const filtered=prepList.filter(m=>(m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase()))&&(fcls==="all"||m.class===fcls));
  return (
    <div>
      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:220}} placeholder={t.search} value={search} onChange={e=>setSearch(e.target.value)}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
          {hasPerm(mod,"edit")&&<><button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button><button className="btn btn-secondary" onClick={()=>fileRef.current.click()}>{t.importXlsx}</button><input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={doImport}/></>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>
      <div style={{display:"flex",gap:7,marginBottom:10,flexWrap:"wrap"}}>
        <button className={`filter-btn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all}</button>
        {cls.map(c=><button key={c} className={`filter-btn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c}</button>)}
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>{["#",t.code,t.name,t.class,t.unit,lang==="ar"?"مكونات":"Ing.",t.yieldWeight,t.costPerUnit,(hasPerm(mod,"edit")||hasPerm(mod,"delete"))?t.actions:""].filter(Boolean).map((h,i)=><th key={i}>{h}</th>)}</tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={9} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :filtered.map((m,i)=>{ const {costPerUnit,yieldKg}=calcPrepCost(m); return(
              <tr key={m.id}>
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>{m.name}</td>
                <td><span className="badge badge-cls">{m.class||"—"}</span></td>
                <td><span className={`badge badge-${m.unit}`}>{unitLbl(m.unit,t)}</span></td>
                <td style={{color:C.muted}}>{m.ingredients?.length||0}</td>
                <td style={{color:C.muted}}>{yieldKg.toFixed(3)}</td>
                <td style={{color:C.accent,fontWeight:700}}>{costPerUnit.toFixed(4)}</td>
                {<td><div style={{display:"flex",gap:5}}>
                <button style={{background:"#0a2a3a",color:"#38bdf8",padding:"4px 11px",fontSize:12,border:"1px solid #0ea5e933",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontWeight:600}} onClick={()=>setViewItem(m)}>{lang==="ar"?"عرض":"View"}</button>
                {hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>doEdit(m)}>{t.edit}</button>}
                {hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>setDelId(m.id)}>{t.delete}</button>}
              </div></td>}
              </tr>
            );})}
          </tbody>
        </table>
      </div></div>

      {/* VIEW MODAL */}
      {viewItem&&(()=>{
        const {costPerUnit,yieldKg}=calcPrepCost(viewItem);
        const totalCost=costPerUnit*yieldKg;
        return(
          <div className="overlay" onClick={e=>e.target===e.currentTarget&&setViewItem(null)}>
            <div className="modal modal-lg" style={{maxWidth:820}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
                <div>
                  <div style={{fontWeight:800,fontSize:16,color:DARK.accent}}>{viewItem.name}</div>
                  <div style={{fontSize:12,color:DARK.muted,marginTop:2}}>{viewItem.code} · {viewItem.class||"—"} · {unitLbl(viewItem.unit,t)}</div>
                </div>
                <button onClick={()=>setViewItem(null)} style={{background:"transparent",border:"none",color:DARK.muted,fontSize:20,cursor:"pointer"}}>✕</button>
              </div>

              {/* summary cards */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:18}}>
                {[
                  {l:lang==="ar"?"عدد المكونات":"Ingredients", v:viewItem.ingredients?.length||0, c:DARK.blue},
                  {l:lang==="ar"?"وزن الناتج":"Yield Weight",  v:yieldKg.toFixed(3)+" "+unitLbl(viewItem.unit,t), c:DARK.green},
                  {l:lang==="ar"?"تكلفة/وحدة":"Cost/Unit",    v:costPerUnit.toFixed(4), c:DARK.accent},
                ].map((s,i)=>(
                  <div key={i} style={{background:DARK.surface,borderRadius:10,padding:"12px 16px",border:"1px solid "+DARK.border}}>
                    <div style={{fontSize:11,color:DARK.muted,marginBottom:4,textTransform:"uppercase",letterSpacing:".05em"}}>{s.l}</div>
                    <div style={{fontSize:18,fontWeight:800,color:s.c}}>{s.v}</div>
                  </div>
                ))}
              </div>

              {/* ingredients table */}
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead>
                    <tr style={{background:DARK.surface}}>
                      {["#",
                        lang==="ar"?"المادة":"Material",
                        lang==="ar"?"الكمية الصافية":"Net Qty",
                        lang==="ar"?"الهدر %":"Wastage %",
                        lang==="ar"?"الكمية الإجمالية":"Gross Qty",
                        lang==="ar"?"سعر/وحدة":"Cost/Unit",
                        lang==="ar"?"إجمالي التكلفة":"Total Cost"
                      ].map((h,i)=><th key={i} style={{padding:"10px 13px",textAlign:lang==="ar"?"right":"left",fontSize:10,fontWeight:700,color:DARK.muted,textTransform:"uppercase",letterSpacing:".06em",whiteSpace:"nowrap"}}>{h}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {(viewItem.ingredients||[]).map((ing,i)=>{
                      const raw=rawList.find(r=>String(r.id)===String(ing.rawId));
                      if(!raw) return null;
                      const qty=parseFloat(ing.qty)||0;
                      const waste=(parseFloat(ing.waste)||0)/100;
                      const grossQty=qty; // raw qty (cost basis)
                      const netQty=qty*(1-waste); // after waste
                      const ingCost=(raw.unit==="piece"?grossQty:grossQty/1000)*raw.price;
                      const unit=raw.unit==="kg"?"g":raw.unit==="liter"?"ml":"pcs";
                      return(
                        <tr key={i} style={{borderBottom:"1px solid "+DARK.border+"22"}}>
                          <td style={{padding:"11px 13px",color:DARK.muted,fontSize:11}}>{i+1}</td>
                          <td style={{padding:"11px 13px",fontWeight:600,color:DARK.text}}>
                            {raw.name}
                            <span style={{color:DARK.muted,fontSize:11,marginRight:6}}> ({raw.code})</span>
                          </td>
                          <td style={{padding:"11px 13px",color:DARK.text}}>{qty.toFixed(0)} {unit}</td>
                          <td style={{padding:"11px 13px"}}>
                            {ing.waste>0
                              ? <span style={{background:DARK.yellow+"22",color:DARK.yellow,padding:"2px 8px",borderRadius:20,fontSize:12,fontWeight:700}}>{ing.waste}%</span>
                              : <span style={{color:DARK.muted}}>0%</span>
                            }
                          </td>
                          <td style={{padding:"11px 13px",color:DARK.green,fontWeight:600}}>{netQty.toFixed(0)} {unit}</td>
                          <td style={{padding:"11px 13px",color:DARK.muted}}>{raw.price.toFixed(2)}/{unitLbl(raw.unit,t)}</td>
                          <td style={{padding:"11px 13px",color:DARK.accent,fontWeight:700}}>{ingCost.toFixed(4)}</td>
                        </tr>
                      );
                    })}
                    {/* totals row */}
                    <tr style={{background:DARK.surface,borderTop:"2px solid "+DARK.border}}>
                      <td colSpan={2} style={{padding:"12px 13px",fontWeight:800,color:DARK.text,fontSize:13}}>
                        {lang==="ar"?"الإجمالي":"Total"}
                      </td>
                      <td style={{padding:"12px 13px",fontWeight:700,color:DARK.text}}>
                        {((viewItem.ingredients||[]).reduce((a,i)=>a+(parseFloat(i.qty)||0),0)).toFixed(0)} g
                      </td>
                      <td/>
                      <td style={{padding:"12px 13px",fontWeight:700,color:DARK.green}}>
                        {((viewItem.ingredients||[]).reduce((a,ing)=>{
                          const w=(parseFloat(ing.waste)||0)/100;
                          const q=parseFloat(ing.qty)||0;
                          return a+q*(1-w);
                        },0)).toFixed(0)} g
                      </td>
                      <td/>
                      <td style={{padding:"12px 13px",fontWeight:800,color:DARK.accent,fontSize:14}}>{totalCost.toFixed(4)}</td>
                    </tr>
                    {/* cost per unit row */}
                    <tr style={{background:DARK.accent+"10"}}>
                      <td colSpan={6} style={{padding:"10px 13px",fontWeight:700,color:DARK.text,fontSize:13,textAlign:lang==="ar"?"right":"left"}}>
                        {lang==="ar"?"تكلفة الكيلو الواحد":"Cost per "+unitLbl(viewItem.unit,t)}
                      </td>
                      <td style={{padding:"10px 13px",fontWeight:900,color:DARK.accent,fontSize:16}}>{costPerUnit.toFixed(4)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>

              <div style={{display:"flex",justifyContent:"flex-end",marginTop:16}}>
                <button className="btn btn-secondary" onClick={()=>setViewItem(null)}>{t.cancel}</button>
              </div>
            </div>
          </div>
        );
      })()}

      {usageModal&&(
        <div className="overlay" onClick={e=>e.target===e.currentTarget&&setUsageModal(null)}>
          <div className="modal" style={{maxWidth:500}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div>
                <div style={{fontWeight:800,fontSize:15,color:DARK.accent}}>{usageModal.item.name}</div>
                <div style={{fontSize:12,color:DARK.muted,marginTop:2}}>{t.relatedProducts} ({usageModal.list.length})</div>
              </div>
              <button onClick={()=>setUsageModal(null)} style={{background:"transparent",border:"none",color:DARK.muted,fontSize:20,cursor:"pointer"}}>✕</button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {usageModal.list.map((item,i)=>(
                <div key={item.id} style={{background:DARK.surface,borderRadius:9,padding:"10px 14px",border:"1px solid "+DARK.border,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <code style={{background:DARK.bg,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8",marginLeft:8}}>{item.code}</code>
                    <span style={{fontWeight:600,fontSize:13}}>{item.name}</span>
                  </div>
                  <span className="badge badge-cls">{item.class||"—"}</span>
                </div>
              ))}
            </div>
            <div style={{display:"flex",justifyContent:"flex-end",marginTop:14}}>
              <button className="btn btn-secondary" onClick={()=>setUsageModal(null)}>{t.cancel}</button>
            </div>
          </div>
        </div>
      )}
      {showForm&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&reset()}><div className="modal modal-lg">
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.edit:t.add} — {t.prepItem}</h2>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:12}}>
          <div style={{gridColumn:"1/3"}}><label className="lbl">{t.name}</label><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">{t.class}</label><select value={form.class} onChange={e=>setForm({...form,class:e.target.value})}><option value="">—</option>{cls.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
          <div><label className="lbl">{t.unit}</label><select value={form.unit} onChange={e=>setForm({...form,unit:e.target.value})}><option value="kg">{t.kg}</option><option value="liter">{t.liter}</option><option value="piece">{t.piece}</option></select></div>
          <div style={{gridColumn:"2/4"}}><label className="lbl">{t.yieldWeight} {t.yieldAuto}</label><input type="number" min="0" step="0.001" value={form.yieldOverride} onChange={e=>setForm({...form,yieldOverride:e.target.value})} placeholder={live.yieldKg.toFixed(3)}/></div>
        </div>
        <div className="divider"/>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <span style={{fontWeight:700,fontSize:13}}>{t.ingredients}</span>
          <button className="btn btn-secondary" style={{padding:"5px 12px",fontSize:12}} onClick={addI}>+ {t.addIngredient}</button>
        </div>
        {form.ingredients.map(ing=>(
          <IngRow key={ing.id} ing={ing} rawList={rawList} prepList={prepList} lang={lang} t={t} C={C}
            onUpdate={(field,val)=>updI(ing.id,field,val)}
            onRemove={()=>remI(ing.id)}
          />
        ))}
        <div style={{background:C.bg,borderRadius:8,padding:"10px 13px",marginTop:8,display:"flex",gap:16,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:C.muted}}>{t.yieldWeight}: <strong style={{color:C.green}}>{live.yieldKg.toFixed(3)}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>{t.costPerUnit}: <strong style={{color:C.accent}}>{live.costPerUnit.toFixed(4)}</strong></span>
        </div>
        <div style={{display:"flex",gap:8,marginTop:16}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
      </div></div>}
      {delId&&<DelModal t={t} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// PRODUCTS
// ═══════════════════════════════════════════════════════════════
function ProductsTab({t,lang,C=DARK,prodList,setProdList,rawList,prepList,classes,calcPrepCost,calcProductCost,showToast,hasPerm,mod}) {
  const [search,setSearch]=useState(""); const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",class:"",sellingPrice:"",stdCost:"",ingredients:[]}); const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null); const fileRef=useRef();
  const [ingSearch,setIngSearch]=useState("");
  const cls=[...(classes.raw||[]),...(classes.prep||[])];
  const blank=()=>({id:Date.now()+Math.random(),source:"raw",srcId:"",qty:"",waste:"0"});
  const reset=()=>{ setForm({name:"",class:"",sellingPrice:"",stdCost:"",ingredients:[]}); setErrs({}); setShowForm(false); setEditId(null); };
  const ok=()=>{ const e={}; if(!form.name.trim())e.name=t.required; setErrs(e); return !Object.keys(e).length; };
  const save=()=>{ if(!ok()) return; const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US"); const ings=form.ingredients.filter(i=>i.srcId&&parseFloat(i.qty)>0); if(editId!==null) setProdList(p=>p.map(m=>m.id===editId?{...m,name:form.name.trim(),class:form.class,sellingPrice:parseFloat(form.sellingPrice)||0,ingredients:ings,lastUpdated:now}:m)); else { const code=genCode("Prod",prodList); setProdList(p=>[...p,{id:Date.now(),code,name:form.name.trim(),class:form.class,sellingPrice:parseFloat(form.sellingPrice)||0,ingredients:ings,lastUpdated:now}]); } reset(); showToast(t.savedOk); };
  const doEdit=m=>{ setForm({name:m.name,class:m.class||"",sellingPrice:String(m.sellingPrice||""),stdCost:String(m.stdCost||""),ingredients:m.ingredients||[]}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setProdList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };
  const addI=()=>setForm(f=>({...f,ingredients:[...f.ingredients,blank()]}));
  const remI=id=>setForm(f=>({...f,ingredients:f.ingredients.filter(i=>i.id!==id)}));
  const updI=(id,k,v)=>setForm(f=>({...f,ingredients:f.ingredients.map(i=>i.id===id?{...i,[k]:v}:i)}));
  const live=calcProductCost({ingredients:form.ingredients,sellingPrice:form.sellingPrice});
  const srcOpts=s=>s==="raw"?rawList:prepList;
  const doExport=()=>{ const rows=prodList.map((p,i)=>{ const {totalCost,margin}=calcProductCost(p); return {"#":i+1,[t.code]:p.code,[t.name]:p.name,[t.class]:p.class||"",[t.totalCost]:totalCost.toFixed(4),[t.sellingPrice]:p.sellingPrice||0,[t.margin]:margin.toFixed(2)+"%"}; }); const ws=XLSX.utils.json_to_sheet(rows); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,t.products); XLSX.writeFile(wb,`products_${Date.now()}.xlsx`); };
  const doImport=e=>{ const file=e.target.files[0]; if(!file) return; const r=new FileReader(); r.onload=ev=>{ try{ const wb=XLSX.read(ev.target.result,{type:"binary"}); const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]); let n=0; const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US"); setProdList(prev=>{ const u=[...prev]; rows.forEach(row=>{ const rn=String(row[t.name]||row.Name||row["الاسم"]||"").trim(); const rc=String(row[t.code]||row.Code||row["الكود"]||"").trim(); const rsp=parseFloat(row[t.sellingPrice]||row["Selling Price"]||0); let i=u.findIndex(m=>m.name.toLowerCase()===rn.toLowerCase()); if(i===-1) i=u.findIndex(m=>m.code===rc); if(i!==-1){u[i]={...u[i],lastUpdated:now}; if(rsp>0)u[i].sellingPrice=rsp; n++;} }); return u; }); showToast(n>0?`${t.importedOk} ${n} ${t.importedItems}`:t.noMatch,n>0?"success":"warning"); }catch{ showToast("error","error"); } }; r.readAsBinaryString(file); e.target.value=""; };
  const filtered=prodList.filter(m=>(m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase()))&&(fcls==="all"||m.class===fcls));
  return (
    <div>
      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:220}} placeholder={t.search} value={search} onChange={e=>setSearch(e.target.value)}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
          {hasPerm(mod,"edit")&&<><button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button><button className="btn btn-secondary" onClick={()=>fileRef.current.click()}>{t.importXlsx}</button><input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={doImport}/></>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>
      <div style={{display:"flex",gap:7,marginBottom:10,flexWrap:"wrap"}}>
        <button className={`filter-btn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all}</button>
        {cls.map(c=><button key={c} className={`filter-btn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c}</button>)}
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>{["#",t.code,t.productName,t.class,lang==="ar"?"مكونات":"Ing.",t.totalCost,t.sellingPrice,t.margin,(hasPerm(mod,"edit")||hasPerm(mod,"delete"))?t.actions:""].filter(Boolean).map((h,i)=><th key={i}>{h}</th>)}</tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={9} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :filtered.map((m,i)=>{ const {totalCost,margin}=calcProductCost(m); return(
              <tr key={m.id}>
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>{m.name}</td>
                <td><span className="badge badge-cls">{m.class||"—"}</span></td>
                <td style={{color:C.muted}}>{m.ingredients?.length||0}</td>
                <td style={{color:"#f87171",fontWeight:600}}>{totalCost.toFixed(2)}</td>
                <td style={{color:DARK.muted,fontWeight:500}}>{m.stdCost>0?m.stdCost.toFixed(2):"—"}</td>
                <td>{m.stdCost>0?(()=>{const v=totalCost-m.stdCost;return <span style={{color:v>0?"#f87171":"#4ade80",fontWeight:700}}>{v>0?"+":""}{v.toFixed(2)}</span>;})():<span style={{color:DARK.muted}}>—</span>}</td>
                <td>{m.stdCost>0?(()=>{const vp=((totalCost-m.stdCost)/m.stdCost)*100;return <span style={{color:vp>0?"#f87171":vp<0?"#4ade80":DARK.muted,fontWeight:700,background:vp>0?"#f8717115":"#4ade8015",padding:"2px 7px",borderRadius:20}}>{vp>0?"+":""}{vp.toFixed(1)}%</span>;})():<span style={{color:DARK.muted}}>—</span>}</td>
                <td style={{color:"#4ade80",fontWeight:600}}>{parseFloat(m.sellingPrice||0).toFixed(2)}</td>
                <td><span style={{color:margin>30?"#4ade80":margin>15?"#fbbf24":"#f87171",fontWeight:700}}>{margin.toFixed(1)}%</span></td>
                {(hasPerm(mod,"edit")||hasPerm(mod,"delete"))&&<td><div style={{display:"flex",gap:5}}>{hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>doEdit(m)}>{t.edit}</button>}{hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>setDelId(m.id)}>{t.delete}</button>}</div></td>}
              </tr>
            );})}
          </tbody>
        </table>
      </div></div>
      {showForm&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&reset()}><div className="modal modal-lg">
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.edit:t.add} — {t.products}</h2>
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:10,marginBottom:12}}>
          <div><label className="lbl">{t.productName}</label><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">{t.class}</label><select value={form.class} onChange={e=>setForm({...form,class:e.target.value})}><option value="">—</option>{cls.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
          <div><label className="lbl">{t.sellingPrice}</label><input type="number" min="0" step="0.01" value={form.sellingPrice} onChange={e=>setForm({...form,sellingPrice:e.target.value})} placeholder="0.00"/></div>
          <div><label className="lbl">{t.stdCost}</label><input type="number" min="0" step="0.01" value={form.stdCost||""} onChange={e=>setForm({...form,stdCost:e.target.value})} placeholder="0.00"/></div>
        </div>
        <div className="divider"/>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <span style={{fontWeight:700,fontSize:13}}>{t.ingredients}</span>
          <button className="btn btn-secondary" style={{padding:"5px 12px",fontSize:12}} onClick={addI}>+ {t.addIngredient}</button>
        </div>
        {form.ingredients.map(ing=>(
          <IngRowProd key={ing.id} ing={ing} rawList={rawList} prepList={prepList} lang={lang} t={t} C={C}
            onUpdate={(field,val)=>updI(ing.id,field,val)}
            onRemove={()=>remI(ing.id)}
          />
        ))}
        <div style={{background:C.bg,borderRadius:8,padding:"10px 13px",marginTop:8,display:"flex",gap:16,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:C.muted}}>{t.totalCost}: <strong style={{color:C.red}}>{live.totalCost.toFixed(4)}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>{t.sellingPrice}: <strong style={{color:C.green}}>{parseFloat(form.sellingPrice||0).toFixed(2)}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>{t.margin}: <strong style={{color:live.margin>30?C.green:live.margin>15?C.yellow:C.red}}>{live.margin.toFixed(1)}%</strong></span>
        </div>
        <div style={{display:"flex",gap:8,marginTop:16}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
      </div></div>}
      {delId&&<DelModal t={t} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// CLASSES
// ═══════════════════════════════════════════════════════════════
function ClassesTab({t,lang,C=DARK,classes,setClasses,showToast,hasPerm,mod}) {
  const [et,setEt]=useState(null); const [ei,setEi]=useState(null); const [val,setVal]=useState(""); const [err,setErr]=useState("");
  const secs=[{key:"raw",label:t.rawMat,cat:t.rawMatCategory},{key:"prep",label:t.prepItem,cat:t.prepCategory}];
  const startAdd=k=>{ setEt(k); setEi(null); setVal(""); setErr(""); };
  const startEdit=(k,i)=>{ setEt(k); setEi(i); setVal(classes[k][i]); setErr(""); };
  const cancel=()=>{ setEt(null); setEi(null); setVal(""); setErr(""); };
  const save=()=>{ if(!val.trim()){setErr(t.required);return;} const list=[...(classes[et]||[])]; if(ei!==null)list[ei]=val.trim(); else list.push(val.trim()); setClasses(c=>({...c,[et]:list})); showToast(t.savedOk); cancel(); };
  const del=(k,i)=>{ setClasses(c=>({...c,[k]:c[k].filter((_,idx)=>idx!==i)})); showToast(t.deletedOk,"error"); };
  return (
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:16}}>
      {secs.map(sec=>(
        <div key={sec.key} className="card" style={{padding:18}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div><div style={{fontWeight:700,fontSize:14}}>{sec.label}</div><div style={{fontSize:11,color:C.muted}}>{t.classFor}: {sec.cat}</div></div>
            {hasPerm(mod,"edit")&&<button className="btn btn-primary" style={{padding:"6px 12px",fontSize:12}} onClick={()=>startAdd(sec.key)}>+ {t.addClass}</button>}
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:7}}>
            {(classes[sec.key]||[]).map((cls,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:C.surface,padding:"9px 12px",borderRadius:8,border:`1px solid ${C.border}`}}>
                {et===sec.key&&ei===i
                  ?<div style={{display:"flex",gap:7,flex:1}}><input value={val} onChange={e=>setVal(e.target.value)} style={{flex:1}}/><button className="btn btn-primary" style={{padding:"4px 11px",fontSize:12}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{padding:"4px 9px",fontSize:12}} onClick={cancel}>x</button></div>
                  :<><span className="badge badge-cls" style={{fontSize:12}}>{cls}</span><div style={{display:"flex",gap:5}}>{hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>startEdit(sec.key,i)}>{t.edit}</button>}{hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>del(sec.key,i)}>{t.delete}</button>}</div></>
                }
              </div>
            ))}
            {(classes[sec.key]||[]).length===0&&<div style={{color:C.muted,textAlign:"center",padding:"16px",fontSize:13}}>{t.noData}</div>}
          </div>
          {et===sec.key&&ei===null&&(
            <div style={{marginTop:12,display:"flex",gap:7}}>
              <input value={val} onChange={e=>setVal(e.target.value)} placeholder={t.className} style={{flex:1}}/>
              <button className="btn btn-primary" style={{padding:"7px 14px",fontSize:13}} onClick={save}>{t.save}</button>
              <button className="btn btn-secondary" style={{padding:"7px 10px",fontSize:13}} onClick={cancel}>x</button>
            </div>
          )}
          {err&&<div className="err" style={{marginTop:5}}>{err}</div>}
        </div>
      ))}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// USERS MANAGEMENT
// ═══════════════════════════════════════════════════════════════
function UsersTab({t,lang,C=DARK,users,setUsers,showToast,currentUserId}) {
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({id:"",name:"",pin:"",role:"user",perms:{raw:{view:true,edit:false,delete:false},prep:{view:true,edit:false,delete:false},products:{view:true,edit:false,delete:false},classes:{view:false,edit:false,delete:false}}});
  const [errs,setErrs]=useState({}); const [delId,setDelId]=useState(null);
  const mods=["raw","prep","products","classes"];
  const modLabel={raw:t.rawMat,prep:t.prepItem,products:t.products,classes:t.classes};
  const reset=()=>{ setForm({id:"",name:"",pin:"",role:"user",perms:{raw:{view:true,edit:false,delete:false},prep:{view:true,edit:false,delete:false},products:{view:true,edit:false,delete:false},classes:{view:false,edit:false,delete:false}}}); setErrs({}); setShowForm(false); setEditId(null); };
  const ok=()=>{ const e={}; if(!form.name.trim())e.name=t.required; if(!editId&&!form.id.trim())e.id=t.required; if(!editId&&!form.pin.trim())e.pin=t.required; if(form.pin&&form.pin.length!==5)e.pin=t.pinHint; if(!editId&&users.find(u=>u.id===form.id.trim()))e.id=lang==="ar"?"رقم المستخدم موجود مسبقاً":"User ID already exists"; setErrs(e); return !Object.keys(e).length; };
  const save=()=>{ if(!ok()) return; if(editId){ setUsers(p=>p.map(u=>u.id===editId?{...u,name:form.name.trim(),role:form.role,perms:form.perms,...(form.pin?{pin:form.pin}:{})}:u)); } else { setUsers(p=>[...p,{id:form.id.trim(),name:form.name.trim(),pin:form.pin,role:form.role,lang:"ar",perms:form.perms}]); } reset(); showToast(t.savedOk); };
  const doEdit=u=>{ setForm({id:u.id,name:u.name,pin:"",role:u.role,perms:u.perms||{raw:{view:true,edit:false,delete:false},prep:{view:true,edit:false,delete:false},products:{view:true,edit:false,delete:false},classes:{view:false,edit:false,delete:false}}}); setEditId(u.id); setShowForm(true); };
  const doDelete=id=>{ setUsers(p=>p.filter(u=>u.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };
  const togglePerm=(mod,perm)=>setForm(f=>({...f,perms:{...f.perms,[mod]:{...f.perms[mod],[perm]:!f.perms[mod]?.[perm]}}}));
  return (
    <div>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}>
        <button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.addUser}</button>
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>{[t.userIdLabel,t.userName,t.userRole,t.permissions,t.actions].map((h,i)=><th key={i}>{h}</th>)}</tr></thead>
          <tbody>
            {users.map(u=>(
              <tr key={u.id}>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{u.id}</code></td>
                <td style={{fontWeight:600}}>{u.name}</td>
                <td><span style={{background:u.role==="admin"?C.accent+"22":"#1a1a33",color:u.role==="admin"?C.accent:"#94a3b8",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700}}>{u.role==="admin"?t.admin:t.user}</span></td>
                <td style={{fontSize:11,color:C.muted}}>{u.role==="admin"?lang==="ar"?"كل الصلاحيات":"All permissions":mods.filter(m=>u.perms?.[m]?.view).map(m=>modLabel[m]).join(", ")||"—"}</td>
                <td><div style={{display:"flex",gap:5}}><button className="btn-sm-e" onClick={()=>doEdit(u)}>{t.edit}</button>{u.id!==currentUserId&&<button className="btn-sm-d" onClick={()=>setDelId(u.id)}>{t.delete}</button>}</div></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div></div>
      {showForm&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&reset()}><div className="modal modal-lg">
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.editUser:t.addUser}</h2>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:14}}>
          {!editId&&<div><label className="lbl">{t.userIdLabel}</label><input value={form.id} onChange={e=>setForm({...form,id:e.target.value})} placeholder="1002"/>{errs.id&&<div className="err">{errs.id}</div>}</div>}
          <div style={editId?{gridColumn:"1/3"}:{}}><label className="lbl">{t.userName}</label><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">{editId?t.newPin:t.pin}</label><input type="password" maxLength={5} value={form.pin} onChange={e=>setForm({...form,pin:e.target.value.replace(/\D/g,"")})} placeholder={editId?lang==="ar"?"اتركه فارغاً للإبقاء":"Leave blank to keep":"•••••"}/>{errs.pin&&<div className="err">{errs.pin}</div>}</div>
          <div><label className="lbl">{t.userRole}</label><select value={form.role} onChange={e=>setForm({...form,role:e.target.value})}><option value="user">{t.user}</option><option value="admin">{t.admin}</option></select></div>
        </div>
        {form.role!=="admin"&&(
          <>
            <div className="divider"/>
            <div style={{fontWeight:700,fontSize:13,marginBottom:10}}>{t.permissions}</div>
            {mods.map(mod=>(
              <div key={mod} className="perm-box">
                <span style={{fontWeight:600,fontSize:13,minWidth:120,color:C.text}}>{modLabel[mod]}</span>
                {["view","edit","delete"].map(perm=>(
                  <label key={perm} style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",fontSize:12,color:C.muted}}>
                    <input type="checkbox" checked={!!form.perms[mod]?.[perm]} onChange={()=>togglePerm(mod,perm)}/>
                    {perm==="view"?t.permView:perm==="edit"?t.permEdit:t.permDelete}
                  </label>
                ))}
              </div>
            ))}
          </>
        )}
        <div style={{display:"flex",gap:8,marginTop:16}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
      </div></div>}
      {delId&&<DelModal t={t} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// DELETE MODAL
// ═══════════════════════════════════════════════════════════════
function DelModal({t,C=DARK,onOk,onCancel}) {
  return (
    <div className="overlay">
      <div className="modal" style={{maxWidth:300,textAlign:"center"}}>
        <p style={{marginBottom:20,color:C.text,fontSize:14}}>{t.confirmDelete}</p>
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          <button className="btn-sm-d" style={{padding:"8px 20px",fontSize:13}} onClick={onOk}>{t.delete}</button>
          <button className="btn btn-secondary" style={{padding:"8px 20px",fontSize:13}} onClick={onCancel}>{t.cancel}</button>
        </div>
      </div>
    </div>
  );
}
