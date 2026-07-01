import React, { useState, useEffect, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { loadFromCloud, saveToCloud } from "./supabase.js";

// ═══════════════════════════════════════════════════════════════
// TRANSLATIONS
// ═══════════════════════════════════════════════════════════════
const T = {
  ar: {
    dir:"rtl", font:"'Cairo',sans-serif",
    appName:"TALA COSTING",
    login:"تسجيل الدخول", userId:"رقم المستخدم", pin:"PIN (5 أرقام)", loginBtn:"دخول", loginErr:"رقم المستخدم أو PIN غير صحيح",
    dashboard:"لوحة التحليل", rawMat:"المواد الخام", prepItem:"المواد شبه المصنعة", products:"المنتجات",
    modifiers:"الموديفايرز", sales:"المبيعات",
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
    userName:"اسم المستخدم", userRole:"الدور", admin:"مدير", user:"مستخدم",
    permissions:"الصلاحيات", permView:"عرض", permEdit:"تعديل", permDelete:"حذف",
    addUser:"إضافة مستخدم", editUser:"تعديل مستخدم", newPin:"PIN جديد",
    pinHint:"5 أرقام فقط", userIdLabel:"رقم المستخدم",
    show:"عرض", items:"صنف",
    secProducts:"تحليل المنتجات", secPrep:"تحليل المواد شبه المصنعة", secRaw:"تحليل المواد الخام",
    prepHighCost:"أعلى تكلفة/وحدة", prepMostUsed:"الأكثر استخداماً في المنتجات",
    rawHighPrice:"أعلى سعر", rawMostUsed:"الأكثر استخداماً",
    usedIn:"مستخدم في", productsCount:"منتج",
    noPermission:"ليس لديك صلاحية للوصول لهذا القسم",
    stdCost:"التكلفة المعيارية", variance:"الانحراف", variancePct:"نسبة الانحراف",
    topCostDriver:"أعلى مكون تكلفة",
    usedInPrep:"في Prep", usedInProducts:"في المنتجات",
    relatedProducts:"المنتجات المرتبطة", relatedPreps:"Prep المرتبطة",
    varianceReport:"تقرير الانحراف عن المعياري",
    marginDistribution:"توزيع هامش الربح",
    topCostDriversReport:"أعلى مكونات التكلفة",
    totalProducts:"إجمالي المنتجات", totalPrep:"إجمالي Prep", totalRaw:"إجمالي الخام",
    avgMarginLbl:"متوسط الهامش", avgCostLbl:"متوسط التكلفة",
    highMargin:"هامش عالي >30%", midMargin:"هامش متوسط 15-30%", lowMargin:"هامش منخفض <15%",
    noSelling:"بدون سعر بيع",
    productsOverBudget:"منتجات تجاوزت المعياري",
    view:"عرض",
    posSellPrice:"سعر POS (المحل)", aggSellPrice:"سعر AGG (التوصيل)",
    posMargin:"هامش POS", aggMargin:"هامش AGG",
    dashPOS:"لوحة POS", dashAGG:"لوحة AGG",
    showMore:"اظهر المزيد", showLess:"اظهر أقل",
    pageSize:"عدد العناصر",
    modifierName:"اسم الموديفاير", modifierPrice:"السعر الإضافي", modifierCost:"التكلفة",
    modifierMargin:"الهامش", addModifier:"إضافة موديفاير",
    salesModule:"المبيعات", salesDate:"التاريخ", salesQty:"الكمية المباعة",
    salesRevenue:"إجمالي المبيعات", salesActualCost:"التكلفة الفعلية",
    salesStdCost:"التكلفة المعيارية", salesCostPct:"نسبة التكلفة",
    salesChannel:"قناة البيع", addSale:"إضافة مبيعات",
    month:"الشهر", year:"السنة",
    monthlySales:"المبيعات الشهرية",
    grossProfit:"مجمل الربح", grossLoss:"مجمل الخسارة",
    import:"استيراد", export:"تصدير",
    cost:"التكلفة",
  },
  en: {
    dir:"ltr", font:"'DM Sans',sans-serif",
    appName:"TALA COSTING",
    login:"Login", userId:"User ID", pin:"PIN (5 digits)", loginBtn:"Sign In", loginErr:"Invalid User ID or PIN",
    dashboard:"Analytics", rawMat:"Raw Materials", prepItem:"Prep Items", products:"Products",
    modifiers:"Modifiers", sales:"Sales",
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
    prepHighCost:"Highest Cost/Unit", prepMostUsed:"Most Used in Products",
    rawHighPrice:"Highest Price", rawMostUsed:"Most Used",
    usedIn:"Used in", productsCount:"products",
    noPermission:"You don't have permission to access this section",
    stdCost:"Std Cost", variance:"Variance", variancePct:"Var %",
    topCostDriver:"Top Cost Driver",
    usedInPrep:"In Prep", usedInProducts:"In Products",
    relatedProducts:"Related Products", relatedPreps:"Related Preps",
    varianceReport:"Variance vs Standard",
    marginDistribution:"Margin Distribution",
    topCostDriversReport:"Top Cost Drivers",
    totalProducts:"Total Products", totalPrep:"Total Prep", totalRaw:"Total Raw",
    avgMarginLbl:"Avg Margin", avgCostLbl:"Avg Cost",
    highMargin:"High Margin >30%", midMargin:"Mid Margin 15-30%", lowMargin:"Low Margin <15%",
    noSelling:"No Selling Price",
    productsOverBudget:"Products Over Budget",
    view:"View",
    posSellPrice:"POS Price (In-store)", aggSellPrice:"AGG Price (Delivery)",
    posMargin:"POS Margin", aggMargin:"AGG Margin",
    dashPOS:"POS Dashboard", dashAGG:"AGG Dashboard",
    showMore:"Show More", showLess:"Show Less",
    pageSize:"Page Size",
    modifierName:"Modifier Name", modifierPrice:"Add-on Price", modifierCost:"Cost",
    modifierMargin:"Margin", addModifier:"Add Modifier",
    salesModule:"Sales", salesDate:"Date", salesQty:"Qty Sold",
    salesRevenue:"Total Revenue", salesActualCost:"Actual Cost",
    salesStdCost:"Std Cost", salesCostPct:"Cost %",
    salesChannel:"Channel", addSale:"Add Sale",
    month:"Month", year:"Year",
    monthlySales:"Monthly Sales",
    grossProfit:"Gross Profit", grossLoss:"Gross Loss",
    import:"Import", export:"Export",
    cost:"Cost",
  }
};

// ═══════════════════════════════════════════════════════════════
// STORAGE + AUTH
// ═══════════════════════════════════════════════════════════════
const SK = {
  raw:"tc_raw_v1", prep:"tc_prep_v1", products:"tc_prods_v1",
  classes:"tc_cls_v1", users:"tc_users_v1", session:"tc_session_v1",
  modifiers:"tc_mods_v1", sales:"tc_sales_v1",
  monthlyPrices:"tc_monthly_prices_v1"
};
const load = (k,d) => { try { const v=localStorage.getItem(k); return v?JSON.parse(v):d; } catch { return d; } };
const persist = (k,v) => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} };

const defaultUsers = [
  { id:"1001", pin:"11111", name:"Admin", role:"admin", lang:"ar",
    perms:{ raw:{view:true,edit:true,delete:true}, prep:{view:true,edit:true,delete:true},
            products:{view:true,edit:true,delete:true}, classes:{view:true,edit:true,delete:true},
            modifiers:{view:true,edit:true,delete:true}, sales:{view:true,edit:true,delete:true} } }
];
const defaultClasses = {
  raw:["Food Item","Package Item","Cleaning Item"],
  prep:["Sauce","Dough","Mix","Marinade"],
  products:["Main Dish","Beverage","Dessert","Appetizer"],
  modifiers:["Add-on","Upgrade","Side","Extra"]
};

const MONTHS_AR_FULL = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"];
const MONTHS_EN_FULL = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const MONTHS_LIST = ["01","02","03","04","05","06","07","08","09","10","11","12"];
const mLabel = (m, lang) => lang==="ar" ? MONTHS_AR_FULL[parseInt(m)-1] : MONTHS_EN_FULL[parseInt(m)-1];

const genCode = (prefix, existing, allLists=[]) => {
  const allCodes = [...existing.map(i=>i.code), ...allLists.flatMap(l=>l.map(i=>i.code))].filter(Boolean);
  const nums = allCodes.filter(c=>c && c.startsWith(prefix+"-")).map(c=>parseInt(c.replace(prefix+"-",""))||0);
  const next = nums.length ? Math.max(...nums)+1 : 1;
  return `${prefix}-${String(next).padStart(5,"0")}`;
};
const unitLbl = (u,t) => ({kg:t.kg,liter:t.liter,piece:t.piece}[u]||u);

const INPUT_UNITS = {
  kg: [{val:"g",label:"جرام / g"},{val:"kg",label:"كيلو / kg"}],
  liter: [{val:"ml",label:"مل / ml"},{val:"liter",label:"لتر / liter"}],
  piece: [{val:"piece",label:"حبة / piece"}],
};
const toInternal = (qty, inputUnit) => {
  if(inputUnit==="kg") return parseFloat(qty)*1000;
  if(inputUnit==="liter") return parseFloat(qty)*1000;
  return parseFloat(qty);
};
const fromInternal = (qty, inputUnit) => {
  if(inputUnit==="kg") return qty/1000;
  if(inputUnit==="liter") return qty/1000;
  return qty;
};

// ═══════════════════════════════════════════════════════════════
// COLORS
// ═══════════════════════════════════════════════════════════════
const DARK = {
  bg:"#080b14", sidebar:"#0d1020", surface:"#111627", card:"#161c2e",
  border:"#1e2540", accent:"#e8a020", accentDark:"#c4861a",
  text:"#dde3f0", muted:"#5a6585", danger:"#ef4444",
  green:"#22c55e", red:"#ef4444", yellow:"#f59e0b", blue:"#3b82f6",
};
const C = DARK;

// ═══════════════════════════════════════════════════════════════
// APP ROOT
// ═══════════════════════════════════════════════════════════════
export default function App() {
  const [users,setUsers]       = useState(()=>load(SK.users,defaultUsers));
  const [session,setSession]   = useState(()=>load(SK.session,null));
  const [rawList,setRawList]   = useState(()=>load(SK.raw,[]));
  const [prepList,setPrepList] = useState(()=>load(SK.prep,[]));
  const [prodList,setProdList] = useState(()=>load(SK.products,[]));
  const [classes,setClasses]   = useState(()=>load(SK.classes,defaultClasses));
  const [modList,setModList]   = useState(()=>load(SK.modifiers,[]));
  const [salesList,setSalesList]= useState(()=>load(SK.sales,[]));
  const [monthlyPrices,setMonthlyPrices]= useState(()=>load(SK.monthlyPrices,{}));
  const [tab,setTab]           = useState("dashboard");
  const [toast,setToast]       = useState(null);
  const [sideOpen,setSideOpen] = useState(true);
  const [cloudReady,setCloudReady] = useState(false);
  const initDone = useRef(false);

  // ── Load from Supabase on startup ─────────────────────────────
  useEffect(()=>{
    if(initDone.current) return;
    initDone.current = true;
    loadFromCloud().then(d=>{
      if(d.users)         setUsers(d.users);
      if(d.raw?.length)   setRawList(d.raw);
      if(d.prep?.length)  setPrepList(d.prep);
      if(d.products?.length) setProdList(d.products);
      if(d.classes)       setClasses(d.classes);
      if(d.modifiers?.length) setModList(d.modifiers);
      if(d.sales?.length) setSalesList(d.sales);
      if(d.monthlyPrices && Object.keys(d.monthlyPrices).length) setMonthlyPrices(d.monthlyPrices);
      setCloudReady(true);
    }).catch(()=>{ setCloudReady(true); });
  },[]);

  // ── Save to localStorage + Supabase on every change ───────────
  useEffect(()=>{ persist(SK.users,users); if(cloudReady) saveToCloud("users",users); },[users,cloudReady]);
  useEffect(()=>{ persist(SK.raw,rawList); if(cloudReady) saveToCloud("raw",rawList); },[rawList,cloudReady]);
  useEffect(()=>{ persist(SK.prep,prepList); if(cloudReady) saveToCloud("prep",prepList); },[prepList,cloudReady]);
  useEffect(()=>{ persist(SK.products,prodList); if(cloudReady) saveToCloud("products",prodList); },[prodList,cloudReady]);
  useEffect(()=>{ persist(SK.classes,classes); if(cloudReady) saveToCloud("classes",classes); },[classes,cloudReady]);
  useEffect(()=>persist(SK.session,session),[session]);
  useEffect(()=>{ persist(SK.modifiers,modList); if(cloudReady) saveToCloud("modifiers",modList); },[modList,cloudReady]);
  useEffect(()=>{ persist(SK.sales,salesList); if(cloudReady) saveToCloud("sales",salesList); },[salesList,cloudReady]);
  useEffect(()=>{ persist(SK.monthlyPrices,monthlyPrices); if(cloudReady) saveToCloud("monthlyPrices",monthlyPrices); },[monthlyPrices,cloudReady]);

  const showToast = useCallback((msg,type="success")=>{ setToast({msg,type}); setTimeout(()=>setToast(null),3000); },[]);

  const currentUser = session ? users.find(u=>u.id===session.id) : null;
  const lang = currentUser?.lang || "ar";
  const t = T[lang];

  const calcPrepCostFn = (prep, depth=0)=>{
    if(!prep.ingredients?.length||depth>5) return {totalCost:0,yieldKg:0,costPerUnit:0};
    let totalCost=0, yieldG=0;
    const isPiece=prep.unit==="piece";
    prep.ingredients.forEach(ing=>{
      const qty=parseFloat(ing.qty)||0;
      const waste=(parseFloat(ing.waste)||0)/100;
      const netQty=qty*(1-waste);
      if(ing.source==="prep"){
        const subPrep=prepList.find(p=>String(p.id)===String(ing.rawId||ing.srcId||ing.prepId));
        if(!subPrep) return;
        const sub=calcPrepCostFn(subPrep,depth+1);
        totalCost+=(subPrep.unit==="piece"?qty:qty/1000)*sub.costPerUnit;
      } else {
        const raw=rawList.find(r=>String(r.id)===String(ing.rawId));
        if(!raw) return;
        totalCost+=(raw.unit==="piece"?qty:qty/1000)*raw.price;
      }
      yieldG+=netQty;
    });
    const yieldKg=prep.yieldOverride&&parseFloat(prep.yieldOverride)>0
      ?parseFloat(prep.yieldOverride)
      :isPiece ? yieldG : (yieldG/1000);
    return {totalCost, yieldKg, costPerUnit:yieldKg>0?totalCost/yieldKg:0};
  };
  const calcPrepCost = useCallback((prep)=>calcPrepCostFn(prep),[rawList,prepList]);

  const calcProductCost = useCallback((prod)=>{
    if(!prod.ingredients?.length) return {totalCost:0,margin:0,posMargin:0,aggMargin:0};
    let cost=0;
    prod.ingredients.forEach(ing=>{
      const qty2=parseFloat(ing.qty)||0;
      if(ing.source==="raw"){
        const raw=rawList.find(r=>String(r.id)===String(ing.srcId)); if(!raw) return;
        cost+=(raw.unit==="piece"?qty2:qty2/1000)*raw.price;
      } else {
        const prep=prepList.find(p=>String(p.id)===String(ing.srcId)); if(!prep) return;
        const {costPerUnit}=calcPrepCost(prep);
        cost+=(prep.unit==="piece"?qty2:qty2/1000)*costPerUnit;
      }
    });
    const posPrice=parseFloat(prod.posSellPrice||prod.sellingPrice)||0;
    const aggPrice=parseFloat(prod.aggSellPrice||prod.sellingPrice)||0;
    const sp=parseFloat(prod.sellingPrice)||posPrice||aggPrice;
    const margin=sp>0?((sp-cost)/sp)*100:0;
    const posMargin=posPrice>0?((posPrice-cost)/posPrice)*100:0;
    const aggMargin=aggPrice>0?((aggPrice-cost)/aggPrice)*100:0;
    return {totalCost:cost, margin, posMargin, aggMargin};
  },[rawList,prepList,calcPrepCost]);

  const hasPerm = (mod,action) => {
    if(!currentUser) return false;
    if(currentUser.role==="admin") return true;
    const p = currentUser.permissions || currentUser.perms || {};
    return p[mod]?.[action]===true;
  };

  if(!currentUser) return <LoginScreen users={users} onLogin={u=>setSession({id:u.id})} lang={lang} />;

  const navItems = [
    {id:"dashboard", label:t.dashboard},
    {id:"raw",       label:t.rawMat,    perm:"raw"},
    {id:"prep",      label:t.prepItem,  perm:"prep"},
    {id:"products",  label:t.products,  perm:"products"},
    {id:"modifiers", label:t.modifiers, perm:"modifiers"},
    {id:"sales",     label:t.sales,     perm:"sales"},
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
        .btn-danger{background:#2a0f0f;color:#f87171;padding:8px 14px;font-size:13px;border:1px solid #dc262633;border-radius:7px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
        .btn-danger:hover{background:#3a1010}
        .btn-sm-e{background:#0f2a4a;color:#60a5fa;padding:4px 11px;font-size:12px;border:1px solid #1e3a6033;border-radius:6px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
        .btn-sm-e:hover{background:#1e3a6033}
        .btn-sm-d{background:#2a0f0f;color:#f87171;padding:4px 11px;font-size:12px;border:1px solid #dc262633;border-radius:6px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
        .btn-sm-d:hover{background:#dc262622}
        .btn-sm-v{background:#0a2a1a;color:#4ade80;padding:4px 11px;font-size:12px;border:1px solid #16a34a33;border-radius:6px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
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
        .fbtn{background:${DARK.surface};border:1px solid ${DARK.border};color:${DARK.muted};padding:5px 12px;font-size:12px;border-radius:20px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
        .fbtn.active{background:${DARK.accent}20;border-color:${DARK.accent}66;color:${DARK.accent}}
        .fbtn:hover{border-color:#3a4060;color:${DARK.text}}
        .divider{height:1px;background:${DARK.border};margin:14px 0}
        .nav-item{display:flex;align-items:center;padding:9px 14px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;color:${DARK.muted};transition:all .14s;border:none;background:transparent;font-family:inherit;width:100%;text-align:${lang==="ar"?"right":"left"}}
        .nav-item:hover{background:${DARK.surface};color:${DARK.text}}
        .nav-item.active{background:${DARK.accent}18;color:${DARK.accent};border-${lang==="ar"?"right":"left"}:3px solid ${DARK.accent}}
        .section-hd{font-size:14px;font-weight:700;color:${DARK.text};margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid ${DARK.border}}
        .topn{background:${DARK.surface};border:1px solid ${DARK.border};color:${DARK.muted};padding:5px 12px;font-size:12px;border-radius:6px;cursor:pointer;font-family:inherit;font-weight:700;transition:all .14s}
        .topn.active{background:${DARK.accent}20;border-color:${DARK.accent}66;color:${DARK.accent}}
        .kpi-card{background:${DARK.surface};border:1px solid ${DARK.border};border-radius:10px;padding:14px 16px}
        .kpi-label{font-size:10px;color:${DARK.muted};font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px}
        .kpi-value{font-size:22px;font-weight:800;line-height:1}
        .sales-kpi-card{background:${DARK.card};border:1px solid ${DARK.border};border-radius:12px;padding:16px 18px}
        .sales-kpi-title{font-size:11px;color:${DARK.muted};font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px}
        .sales-kpi-main{font-size:26px;font-weight:800;line-height:1;margin-bottom:6px}
        .sales-kpi-sub{display:flex;gap:10px;font-size:12px}
        table{width:100%;border-collapse:collapse}
      `}</style>

      {/* SIDEBAR */}
      <div style={{width:sideOpen?220:0,minWidth:sideOpen?220:0,background:DARK.sidebar,borderRight:`1px solid ${DARK.border}`,display:"flex",flexDirection:"column",transition:"all .2s",overflow:"hidden",position:"sticky",top:0,height:"100vh",flexShrink:0}}>
        <div style={{padding:"20px 16px 12px"}}>
          <div style={{fontWeight:900,fontSize:16,color:C.accent,letterSpacing:"1px"}}>{t.appName}</div>
          <div style={{fontSize:10,color:C.muted,marginTop:2,letterSpacing:".5px"}}>SYSTEM v2.3</div>
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
            <button onClick={()=>setSideOpen(o=>!o)} style={{background:"transparent",border:"none",color:C.muted,cursor:"pointer",fontSize:18,padding:"2px 6px"}}>☰</button>
            <span style={{fontWeight:700,fontSize:14,color:C.text}}>{navItems.find(n=>n.id===tab)?.label||""}</span>
          </div>
          <button className="btn btn-secondary" style={{fontSize:12,padding:"5px 12px"}} onClick={()=>{const u=users.find(x=>x.id===currentUser.id);if(u){const newLang=u.lang==="ar"?"en":"ar";setUsers(p=>p.map(x=>x.id===u.id?{...x,lang:newLang}:x));}}}>
            {t.lang}
          </button>
        </div>
        {/* CONTENT */}
        <div style={{flex:1,overflowY:"auto",padding:"20px"}}>
          {tab==="dashboard" && <DashboardTab t={t} lang={lang} C={C} rawList={rawList} prepList={prepList} prodList={prodList} modList={modList} salesList={salesList} calcPrepCost={calcPrepCost} calcProductCost={calcProductCost}/>}
          {tab==="raw"       && (hasPerm("raw","view")?<RawTab t={t} lang={lang} C={C} rawList={rawList} setRawList={setRawList} classes={classes} prepList={prepList} prodList={prodList} showToast={showToast} hasPerm={hasPerm} mod="raw"/>:<NoPerm t={t}/>)}
          {tab==="prep"      && (hasPerm("prep","view")?<PrepTab t={t} lang={lang} C={C} prepList={prepList} setPrepList={setPrepList} rawList={rawList} prodList={prodList} classes={classes} calcPrepCost={calcPrepCost} showToast={showToast} hasPerm={hasPerm} mod="prep"/>:<NoPerm t={t}/>)}
          {tab==="products"  && (hasPerm("products","view")?<ProductsTab t={t} lang={lang} C={C} prodList={prodList} setProdList={setProdList} rawList={rawList} prepList={prepList} classes={classes} calcPrepCost={calcPrepCost} calcProductCost={calcProductCost} showToast={showToast} hasPerm={hasPerm} mod="products"/>:<NoPerm t={t}/>)}
          {tab==="modifiers" && (hasPerm("modifiers","view")?<ModifiersTab t={t} lang={lang} C={C} rawList={rawList} prepList={prepList} classes={classes} modList={modList} setModList={setModList} calcPrepCost={calcPrepCost} showToast={showToast} hasPerm={hasPerm} mod="modifiers" setClasses={setClasses}/>:<NoPerm t={t}/>)}
          {tab==="sales"     && (hasPerm("sales","view")?<SalesTab t={t} lang={lang} C={C} prodList={prodList} prepList={prepList} modList={modList} rawList={rawList} salesList={salesList} setSalesList={setSalesList} monthlyPrices={monthlyPrices} setMonthlyPrices={setMonthlyPrices} calcProductCost={calcProductCost} showToast={showToast} hasPerm={hasPerm} mod="sales"/>:<NoPerm t={t}/>)}
          {tab==="classes"   && (hasPerm("classes","view")?<ClassesTab t={t} lang={lang} C={C} classes={classes} setClasses={setClasses} showToast={showToast} hasPerm={hasPerm} mod="classes"/>:<NoPerm t={t}/>)}
          {tab==="users"     && currentUser.role==="admin" && <UsersTab t={t} lang={lang} C={C} users={users} setUsers={setUsers} showToast={showToast} currentUserId={currentUser.id}/>}
        </div>
      </div>
      {toast&&<div className={`toast ${toast.type}`}>{toast.msg}</div>}
    </div>
  );
}

function NoPerm({t,C=DARK}) {
  return <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:200,color:C.muted,fontSize:14}}>{t.noPermission}</div>;
}


// ═══════════════════════════════════════════════════════════════
// LOGIN SCREEN
// ═══════════════════════════════════════════════════════════════
function LoginScreen({users,onLogin,lang:initLang}) {
  const [uid,setUid]=useState(""); const [pin,setPin]=useState(""); const [err,setErr]=useState("");
  const [loading,setLoading]=useState(false); const [shake,setShake]=useState(false);
  const [focused,setFocused]=useState(null); const [lang,setLang]=useState(initLang||"ar");
  const t=T[lang]; const isRtl=lang==="ar";
  const pinDots=Array(5).fill(0).map((_,i)=>i<pin.length);
  const handleLogin=()=>{
    setLoading(true);
    setTimeout(()=>{
      const u=users.find(x=>x.id===uid.trim()&&x.pin===pin);
      if(u){onLogin(u);}
      else{setErr(t.loginErr);setShake(true);setTimeout(()=>setShake(false),600);}
      setLoading(false);
    },400);
  };
  return (
    <div dir={isRtl?"rtl":"ltr"} style={{minHeight:"100vh",background:"#060810",fontFamily:isRtl?"'Cairo',sans-serif":"'DM Sans',sans-serif",display:"flex",alignItems:"center",justifyContent:"center",color:"#dde3f0",position:"relative",overflow:"hidden"}}>
      <style>{`@keyframes fadeUp{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}@keyframes shakeX{0%,100%{transform:translateX(0)}15%{transform:translateX(-10px)}30%{transform:translateX(10px)}45%{transform:translateX(-8px)}60%{transform:translateX(8px)}75%{transform:translateX(-4px)}90%{transform:translateX(4px)}}@keyframes pulse{0%,100%{opacity:.4;transform:scale(1)}50%{opacity:.7;transform:scale(1.04)}}@keyframes dotPop{from{transform:scale(0)}to{transform:scale(1)}}@keyframes loadBar{from{width:0}to{width:100%}}.lf-wrap{animation:fadeUp .55s ease both}.lf-card{animation:fadeUp .55s .1s ease both}.lf-shake{animation:shakeX .55s ease}.lf-input-wrap{position:relative;margin-bottom:22px}.lf-input{width:100%;background:transparent;border:none;border-bottom:2px solid #1e2540;color:#dde3f0;padding:14px 0 10px;font-size:15px;outline:none;font-family:inherit;transition:border-color .2s;letter-spacing:.5px;}.lf-input:focus{border-bottom-color:#e8a020}.lf-label{position:absolute;top:14px;${isRtl?"right":"left"}:0;font-size:12px;font-weight:700;color:#3a4a6a;text-transform:uppercase;letter-spacing:.1em;transition:all .18s;pointer-events:none;}.lf-input:focus ~ .lf-label,.lf-input:not(:placeholder-shown) ~ .lf-label{top:-4px;font-size:10px;color:#e8a020;letter-spacing:.12em;}.lf-btn{width:100%;background:linear-gradient(135deg,#e8a020,#c4861a);color:#060810;border:none;border-radius:10px;padding:14px;font-size:14px;font-weight:800;font-family:inherit;cursor:pointer;letter-spacing:.08em;text-transform:uppercase;transition:all .2s;position:relative;overflow:hidden;}.lf-btn:hover{filter:brightness(1.1);box-shadow:0 8px 32px #e8a02044;transform:translateY(-1px)}.lf-btn-loading{pointer-events:none;opacity:.8}.lf-loadbar{position:absolute;bottom:0;left:0;height:3px;background:rgba(0,0,0,.3);animation:loadBar .8s ease forwards}.lf-err{color:#f87171;font-size:12px;text-align:center;margin-bottom:16px;font-weight:600;animation:fadeUp .2s ease}.lf-lang{background:transparent;border:1px solid #1e2540;color:#5a6585;padding:6px 18px;border-radius:20px;cursor:pointer;font-size:12px;font-weight:700;font-family:inherit;letter-spacing:.06em;transition:all .18s}.lf-lang:hover{border-color:#e8a02066;color:#e8a020}.dot{width:12px;height:12px;border-radius:50%;border:2px solid #3a4a6a;transition:all .2s;display:inline-block;margin:0 4px}.dot.filled{background:#e8a020;border-color:#e8a020;animation:dotPop .15s ease;box-shadow:0 0 10px #e8a02066}`}</style>
      <div style={{position:"absolute",inset:0,overflow:"hidden",pointerEvents:"none"}}>
        <div style={{position:"absolute",inset:0,background:"radial-gradient(ellipse 80% 60% at 50% 0%,#0d1428 0%,#060810 70%)"}}/>
        <svg width="100%" height="100%" style={{position:"absolute",inset:0,opacity:.07}}><defs><pattern id="grid" width="60" height="60" patternUnits="userSpaceOnUse"><path d="M 60 0 L 0 0 0 60" fill="none" stroke="#e8a020" strokeWidth=".8"/></pattern></defs><rect width="100%" height="100%" fill="url(#grid)"/></svg>
        <div style={{position:"absolute",top:"-15%",left:"60%",width:500,height:500,background:"radial-gradient(circle,#e8a02018 0%,transparent 70%)",animation:"pulse 6s ease-in-out infinite"}}/>
      </div>
      <div style={{display:"flex",width:"100%",maxWidth:900,position:"relative",zIndex:1,minHeight:"100vh"}}>
        <div style={{flex:1,flexDirection:"column",justifyContent:"center",padding:"60px 80px",display:"flex"}}>
          <div className="lf-wrap">
            <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:52,fontWeight:800,color:"#dde3f0",lineHeight:1.05,marginBottom:16}}>TALA<br/><span style={{color:"#e8a020"}}>COSTING</span></div>
            <div style={{width:48,height:3,background:"linear-gradient(90deg,#e8a020,transparent)",marginBottom:20,borderRadius:2}}/>
            <div style={{fontSize:14,color:"#5a6585",lineHeight:1.8,maxWidth:320}}>{isRtl?"نظام متكامل لإدارة وتحليل تكاليف المنتجات":"Integrated cost management system"}</div>
          </div>
        </div>
        <div style={{width:440,minWidth:440,display:"flex",flexDirection:"column",justifyContent:"center",padding:"60px 48px",background:"#0a0d1a",borderLeft:"1px solid #1e2540",position:"relative"}}>
          <div style={{position:"absolute",top:0,left:0,width:80,height:3,background:"linear-gradient(90deg,#e8a020,transparent)"}}/>
          <div className="lf-card">
            <div style={{marginBottom:32}}>
              <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:22,fontWeight:800,color:"#dde3f0"}}>TALA <span style={{color:"#e8a020"}}>COSTING</span></div>
              <div style={{fontSize:12,color:"#3a4a6a",marginTop:4,fontWeight:600,letterSpacing:".1em",textTransform:"uppercase"}}>{isRtl?"تسجيل الدخول":"Sign in to continue"}</div>
            </div>
            <div className={`lf-input-wrap${shake?" lf-shake":""}`}>
              <input className="lf-input" value={uid} placeholder=" " onChange={e=>setUid(e.target.value)} onFocus={()=>setFocused("uid")} onBlur={()=>setFocused(null)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} style={{textAlign:isRtl?"right":"left"}}/>
              <span className="lf-label">{t.userId}</span>
            </div>
            <div className={`lf-input-wrap${shake?" lf-shake":""}`} style={{marginBottom:10}}>
              <input className="lf-input" type="password" maxLength={5} value={pin} placeholder=" " onChange={e=>setPin(e.target.value.replace(/\D/g,""))} onFocus={()=>setFocused("pin")} onBlur={()=>setFocused(null)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} style={{opacity:0,position:"absolute",width:"100%",cursor:"default"}}/>
              <span className="lf-label" style={{top:focused==="pin"||pin.length>0?"-4px":"14px",fontSize:focused==="pin"||pin.length>0?"10px":"12px",color:focused==="pin"?"#e8a020":"#3a4a6a"}}>{t.pin}</span>
              <div style={{borderBottom:`2px solid ${focused==="pin"?"#e8a020":"#1e2540"}`,padding:"14px 0 14px",display:"flex",gap:10,alignItems:"center",justifyContent:isRtl?"flex-end":"flex-start",cursor:"text",transition:"border-color .2s"}} onClick={()=>document.querySelector('input[type="password"]').focus()}>
                {pinDots.map((filled,i)=><span key={i} className={`dot${filled?" filled":""}`}/>)}
              </div>
            </div>
            {err&&<div className="lf-err">{err}</div>}
            <button className={`lf-btn${loading?" lf-btn-loading":""}`} onClick={handleLogin} style={{marginTop:12}}>
              {loading?<span style={{display:"flex",alignItems:"center",justifyContent:"center",gap:8}}><span style={{width:16,height:16,border:"2px solid #06081088",borderTop:"2px solid #060810",borderRadius:"50%",display:"inline-block",animation:"spin .7s linear infinite"}}/>{isRtl?"جاري...":"Signing in..."}<div className="lf-loadbar"/></span>:t.loginBtn}
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


// ═══════════════════════════════════════════════════════════════
// DASHBOARD — Redesigned with correct card order from PDF
// ═══════════════════════════════════════════════════════════════
function DashboardTab({t,lang,C=DARK,rawList,prepList,prodList,modList,salesList,calcPrepCost,calcProductCost}) {
  const [topN,setTopN]=useState(10);
  const [section,setSection]=useState("all");
  const [relModal,setRelModal]=useState(null);
  const ar=lang==="ar";

  const prodCalc=useMemo(()=>prodList.map(p=>({...p,...calcProductCost(p)})),[prodList,calcProductCost]);
  const prepCalc=useMemo(()=>prepList.map(p=>({...p,...calcPrepCost(p)})),[prepList,calcPrepCost]);

  const withPOS=prodCalc.filter(p=>parseFloat(p.posSellPrice||p.sellingPrice)>0);
  const withAGG=prodCalc.filter(p=>parseFloat(p.aggSellPrice||p.sellingPrice)>0);
  const avgPOSMargin=withPOS.length?withPOS.reduce((a,p)=>a+p.posMargin,0)/withPOS.length:0;
  const avgAGGMargin=withAGG.length?withAGG.reduce((a,p)=>a+p.aggMargin,0)/withAGG.length:0;
  const avgPOSCost=withPOS.length?withPOS.reduce((a,p)=>a+p.totalCost,0)/withPOS.length:0;
  const avgAGGCost=withAGG.length?withAGG.reduce((a,p)=>a+p.totalCost,0)/withAGG.length:0;

  const maxCost=prodCalc.length?Math.max(...prodCalc.map(p=>p.totalCost)):0;
  const minCost=prodCalc.filter(p=>p.totalCost>0).length?Math.min(...prodCalc.filter(p=>p.totalCost>0).map(p=>p.totalCost)):0;
  const withStd=prodCalc.filter(p=>parseFloat(p.stdCost)>0);
  const overBudget=withStd.filter(p=>p.totalCost>parseFloat(p.stdCost));

  const getMargin=(p)=>p.posMargin;
  const highM=prodCalc.filter(p=>getMargin(p)>30&&parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0).length;
  const midM=prodCalc.filter(p=>getMargin(p)>=15&&getMargin(p)<=30&&parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0).length;
  const lowM=prodCalc.filter(p=>getMargin(p)<15&&parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0&&getMargin(p)>0).length;
  const noPrice=prodCalc.filter(p=>!parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)).length;

  // Sales analytics
  const totalRevPOS=salesList.filter(s=>s.channel==="POS").reduce((a,s)=>a+(parseFloat(s.revenuePos||0)||parseFloat(s.revenue||0)),0);
  const totalRevAGG=salesList.filter(s=>s.channel==="AGG").reduce((a,s)=>a+(parseFloat(s.revenueAgg||0)||parseFloat(s.revenue||0)),0);
  const totalRev=salesList.reduce((a,s)=>a+(parseFloat(s.revenuePos||0))+(parseFloat(s.revenueAgg||0)),0) || totalRevPOS+totalRevAGG;
  const totalCostSales=salesList.reduce((a,s)=>{
    const p=prodList.find(x=>x.id===s.productId||x.code===s.productCode);
    if(!p) return a;
    const {totalCost}=calcProductCost(p);
    const qty=(parseFloat(s.qtyPos||0)+parseFloat(s.qtyAgg||0))||parseFloat(s.qty||0);
    return a+(totalCost*qty);
  },0);
  const overallSalesMargin=totalRev>0?((totalRev-totalCostSales)/totalRev)*100:0;
  const grossPL=totalRev-totalCostSales;

  const driverMap={};
  prodList.forEach(prod=>{prod.ingredients?.forEach(ing=>{
    if(ing.source==="raw"){const r=rawList.find(r=>String(r.id)===String(ing.srcId));if(!r)return;driverMap[r.name]=(driverMap[r.name]||0)+(parseFloat(ing.qty)||0)/1000*r.price;}
    else{const p=prepList.find(p=>String(p.id)===String(ing.srcId));if(!p)return;const {costPerUnit}=calcPrepCost(p);driverMap[p.name]=(driverMap[p.name]||0)+(parseFloat(ing.qty)||0)/1000*costPerUnit;}
  });});
  const drivers=Object.entries(driverMap).sort((a,b)=>b[1]-a[1]).slice(0,topN);
  const maxDrv=drivers[0]?.[1]||1;

  const prepUsage=prepCalc.map(p=>({...p,
    usedIn:prodList.filter(pr=>pr.ingredients?.some(i=>i.source==="prep"&&String(i.srcId)===String(p.id))).length,
    relatedProds:prodList.filter(pr=>pr.ingredients?.some(i=>i.source==="prep"&&String(i.srcId)===String(p.id)))
  }));
  const rawUsage=rawList.map(r=>({...r,
    inPrep:prepList.filter(p=>p.ingredients?.some(i=>String(i.rawId)===String(r.id))).length,
    inProd:prodList.filter(p=>p.ingredients?.some(i=>i.source==="raw"&&String(i.srcId)===String(r.id))).length,
    relatedPreps:prepList.filter(p=>p.ingredients?.some(i=>String(i.rawId)===String(r.id))),
    relatedProds:prodList.filter(p=>p.ingredients?.some(i=>i.source==="raw"&&String(i.srcId)===String(r.id))),
  }));
  const atRisk=prodCalc.filter(p=>p.totalCost>0&&getMargin(p)<20&&parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0).sort((a,b)=>getMargin(a)-getMargin(b));

  const Bar=({val,max,color})=>(<div style={{background:C.surface,borderRadius:3,height:6,flex:1}}><div style={{width:Math.min((val/(max||1))*100,100)+"%",height:6,borderRadius:3,background:color,transition:"width .4s"}}/></div>);
  const SecHd=({c,sub})=>(<div style={{marginBottom:14}}><div style={{fontWeight:700,fontSize:14,color:C.text}}>{c}</div>{sub&&<div style={{fontSize:11,color:C.muted,marginTop:2}}>{sub}</div>}<div style={{height:2,background:C.border,marginTop:8}}/></div>);
  const noMsg=<div style={{color:C.muted,fontSize:13,textAlign:"center",padding:"24px 0"}}>{t.noData}</div>;

  const secs=[
    {id:"all",label:ar?"الكل":"All"},
    {id:"classification",label:ar?"تصنيف المنتجات":"Classification"},
    {id:"variance",label:t.varianceReport},
    {id:"costanalysis",label:ar?"تحليل التكلفة":"Cost Analysis"},
    {id:"prep",label:t.secPrep},
    {id:"raw",label:t.secRaw},
    {id:"salesAnalysis",label:ar?"تحليل المبيعات":"Sales Analysis"},
  ];
  const show=s=>section==="all"||section===s;

  return (
    <div style={{maxWidth:1100,margin:"0 auto"}}>
      {/* Section nav */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          {secs.map(s=><button key={s.id} className={`fbtn${section===s.id?" active":""}`} onClick={()=>setSection(s.id)}>{s.label}</button>)}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:12,color:C.muted}}>{t.show}</span>
          {[5,10,20].map(n=><button key={n} className={`topn${topN===n?" active":""}`} onClick={()=>setTopN(n)}>{n}</button>)}
        </div>
      </div>

      {/* === CARD 1: KPI Summary === */}
      {show("costanalysis")&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10,marginBottom:16}}>
        {[
          {label:ar?"إجمالي المنتجات":"Total Products",value:prodList.length,color:C.accent},
          {label:ar?"إجمالي الخام":"Total Raw",value:rawList.length,color:"#a78bfa"},
          {label:ar?"إجمالي Prep":"Total Prep",value:prepList.length,color:C.blue},
          {label:ar?"متوسط الهامش POS":"Avg Margin POS",value:avgPOSMargin.toFixed(1)+"%",color:avgPOSMargin>30?C.green:avgPOSMargin>15?C.yellow:C.red,sub:ar?`متوسط تكلفة: ${avgPOSCost.toFixed(2)}`:`Avg cost: ${avgPOSCost.toFixed(2)}`},
          {label:ar?"متوسط الهامش AGG":"Avg Margin AGG",value:avgAGGMargin.toFixed(1)+"%",color:avgAGGMargin>30?C.green:avgAGGMargin>15?C.yellow:C.red,sub:ar?`متوسط تكلفة: ${avgAGGCost.toFixed(2)}`:`Avg cost: ${avgAGGCost.toFixed(2)}`},
          {label:ar?"أعلى تكلفة":"Max Cost",value:maxCost.toFixed(2),color:"#f87171",sub:ar?`أقل: ${minCost.toFixed(2)}`:`Min: ${minCost.toFixed(2)}`},
          {label:t.productsOverBudget,value:overBudget.length,color:overBudget.length>0?C.red:C.green,sub:withStd.length>0?`/ ${withStd.length} ${ar?"لهم معياري":"with target"}`:""},
        ].map((k,i)=>(
          <div key={i} className="card" style={{padding:"14px 16px"}}>
            <div style={{fontSize:i>=3?28:24,fontWeight:800,color:k.color,lineHeight:1}}>{k.value}</div>
            <div style={{fontSize:11,color:C.muted,marginTop:4}}>{k.label}</div>
            {k.sub&&<div style={{fontSize:10,color:C.muted,marginTop:2}}>{k.sub}</div>}
          </div>
        ))}
      </div>}

      {/* === CARD 2: POS vs AGG Boards === */}
      {show("costanalysis")&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
        {[{ch:"pos",label:t.dashPOS,color:"#22c55e"},{ch:"agg",label:t.dashAGG,color:"#3b82f6"}].map(({ch,label,color})=>{
          const prods=prodCalc.filter(p=>parseFloat(ch==="pos"?(p.posSellPrice||p.sellingPrice):(p.aggSellPrice||p.sellingPrice))>0);
          const avgM=prods.length?prods.reduce((a,p)=>a+(ch==="pos"?p.posMargin:p.aggMargin),0)/prods.length:0;
          const avgCst=prods.length?prods.reduce((a,p)=>a+p.totalCost,0)/prods.length:0;
          const high=prods.filter(p=>(ch==="pos"?p.posMargin:p.aggMargin)>30).length;
          const low=prods.filter(p=>(ch==="pos"?p.posMargin:p.aggMargin)<15&&(ch==="pos"?p.posMargin:p.aggMargin)>0).length;
          return (
            <div key={ch} className="card" style={{padding:"14px 16px",border:`1px solid ${color}22`}}>
              <div style={{fontWeight:700,fontSize:13,color,marginBottom:10,display:"flex",alignItems:"center",gap:6}}>
                <span style={{width:8,height:8,borderRadius:"50%",background:color,display:"inline-block"}}/>{label}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                {[
                  {v:avgM.toFixed(1)+"%",l:t.avgMarginLbl,c:avgM>30?C.green:avgM>15?C.yellow:C.red},
                  {v:avgCst.toFixed(2),l:t.avgCostLbl,c:"#f87171"},
                  {v:high,l:ar?"هامش >30%":"Margin >30%",c:C.green},
                  {v:low,l:ar?"هامش <15%":"Margin <15%",c:C.red},
                ].map((s,i)=>(
                  <div key={i} style={{background:C.surface,borderRadius:7,padding:"8px 10px"}}>
                    <div style={{fontSize:18,fontWeight:800,color:s.c}}>{s.v}</div>
                    <div style={{fontSize:10,color:C.muted}}>{s.l}</div>
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>}

      {/* === CARD 3: Product Classification (Star/Puzzle/Plow Horse/Dog) === */}
      {(show("classification"))&&prodList.length>0&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={ar?"تصنيف المنتجات":"Product Classification"} sub={ar?"حسب الهامش والمبيعات — هامش >30% = عالي":"Based on margin & sales — margin >30% = high"}/>
        {(()=>{
          const salesMap={};
          salesList.forEach(s=>{
            const k=s.productCode;
            if(!salesMap[k]) salesMap[k]={rev:0,qty:0};
            salesMap[k].rev+=(parseFloat(s.revenuePos||0)+parseFloat(s.revenueAgg||0))||parseFloat(s.revenue||0);
            salesMap[k].qty+=(parseFloat(s.qtyPos||0)+parseFloat(s.qtyAgg||0))||parseFloat(s.qty||0);
          });
          const avgRev=Object.keys(salesMap).length>0?Object.values(salesMap).reduce((s,v)=>s+v.rev,0)/Object.keys(salesMap).length:0;
          const classified=prodCalc.map(p=>{
            const margin=getMargin(p);
            const rev=(salesMap[p.code]?.rev||0);
            const highMargin=margin>=30;
            const highSales=salesList.length>0?(rev>=avgRev):true;
            let cat,col;
            if(highMargin&&highSales){cat=ar?"⭐ نجم":"⭐ Star";col="#fbbf24";}
            else if(highMargin&&!highSales){cat=ar?"🔮 لغز":"🔮 Puzzle";col="#a78bfa";}
            else if(!highMargin&&highSales){cat=ar?"🐴 حصان المحراث":"🐴 Plow Horse";col="#3b82f6";}
            else{cat=ar?"🐕 غير مربح":"🐕 Dog";col="#f87171";}
            return {...p,cat,catColor:col,salesRev:rev};
          });
          const groups={};classified.forEach(p=>{if(!groups[p.cat])groups[p.cat]=[];groups[p.cat].push(p);});
          const order=ar?["⭐ نجم","🐴 حصان المحراث","🔮 لغز","🐕 غير مربح"]:["⭐ Star","🐴 Plow Horse","🔮 Puzzle","🐕 Dog"];
          return (
            <div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:12}}>
                {order.map(cat=>{
                  const items=groups[cat]||[];
                  const col=classified.find(p=>p.cat===cat)?.catColor||C.muted;
                  return (
                    <div key={cat} style={{background:C.surface,border:`1px solid ${col}44`,borderRadius:12,padding:"12px 14px"}}>
                      <div style={{fontSize:13,fontWeight:800,color:col,marginBottom:8,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                        <span>{cat}</span>
                        <span onClick={()=>items.length>0&&setRelModal({type:"classification",cat,items,col})} style={{fontSize:12,background:col+"22",color:col,borderRadius:20,padding:"2px 10px",cursor:items.length>0?"pointer":"default",fontWeight:700,border:`1px solid ${col}44`}}>{items.length} {ar?"منتج":"products"}</span>
                      </div>
                      {items.length===0&&<div style={{fontSize:11,color:C.muted}}>—</div>}
                      {items.slice(0,3).map(p=>(
                        <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5,paddingBottom:5,borderBottom:`1px solid ${C.border}22`}}>
                          <span style={{fontSize:11,color:C.text,fontWeight:600,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</span>
                          <span style={{fontSize:11,color:col,fontWeight:700,marginRight:6}}>{getMargin(p).toFixed(1)}%</span>
                        </div>
                      ))}
                      {items.length>3&&<div style={{fontSize:10,color:C.muted,marginTop:4,cursor:"pointer"}} onClick={()=>setRelModal({type:"classification",cat,items,col})}>+{items.length-3} {ar?"أكثر...":"more..."}</div>}
                    </div>
                  );
                })}
              </div>
              {salesList.length===0&&<div style={{marginTop:10,fontSize:11,color:C.muted,background:C.surface,borderRadius:8,padding:"8px 12px"}}>
                {ar?"💡 أضف بيانات مبيعات للحصول على تصنيف دقيق بناءً على الإيرادات الفعلية":"💡 Add sales data for accurate classification based on actual revenue"}
              </div>}
            </div>
          );
        })()}
      </div>}

      {/* === CARD 4: Margin Distribution === */}
      {show("costanalysis")&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.marginDistribution} sub={ar?"توزيع المنتجات حسب هامش POS":"Products by POS margin"}/>
        {prodCalc.length===0?noMsg:<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10,marginBottom:14}}>
            {[{l:t.highMargin,n:highM,col:C.green},{l:t.midMargin,n:midM,col:C.yellow},{l:t.lowMargin,n:lowM,col:C.red},{l:t.noSelling,n:noPrice,col:C.muted}].map((s,i)=>{
              const pct=prodList.length?((s.n/prodList.length)*100).toFixed(0):0;
              return <div key={i} style={{background:C.surface,borderRadius:10,padding:"12px 14px",border:`1px solid ${C.border}`}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}><span style={{fontSize:11,color:C.muted,fontWeight:600}}>{s.l}</span><span style={{fontSize:11,color:s.col,fontWeight:700}}>{pct}%</span></div>
                <div style={{fontSize:22,fontWeight:800,color:s.col}}>{s.n}</div>
                <div style={{background:C.border,borderRadius:3,height:4,marginTop:8}}><div style={{width:pct+"%",height:4,borderRadius:3,background:s.col}}/></div>
              </div>;
            })}
          </div>
          <div style={{display:"flex",height:8,borderRadius:4,overflow:"hidden",gap:1}}>
            {[[highM,C.green],[midM,C.yellow],[lowM,C.red],[noPrice,C.muted]].map(([n,col],i)=><div key={i} style={{flex:n||0,background:col,minWidth:n>0?2:0}}/>)}
          </div>
        </>}
      </div>}

      {/* === CARD 5: Variance Report === */}
      {(show("variance"))&&withStd.length>0&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.varianceReport} sub={ar?"مقارنة التكلفة الفعلية بالمعيارية":"Actual vs Standard Cost"}/>
        <div style={{overflowX:"auto"}}>
          <table style={{borderCollapse:"collapse"}}>
            <thead><tr>{[t.productName,t.totalCost,t.stdCost,t.variance,t.variancePct,t.topCostDriver].map((h,i)=><th key={i} style={{padding:"9px 12px",textAlign:ar?"right":"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",background:C.surface,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>{[...withStd].sort((a,b)=>(b.totalCost-parseFloat(b.stdCost))-(a.totalCost-parseFloat(a.stdCost))).slice(0,topN).map((p,i)=>{
              const v=p.totalCost-parseFloat(p.stdCost); const vp=(v/parseFloat(p.stdCost))*100;
              let topDrv="—",topC2=0;
              p.ingredients?.forEach(ing=>{let cost=0,name="";if(ing.source==="raw"){const r=rawList.find(r=>String(r.id)===String(ing.srcId));if(r){cost=(parseFloat(ing.qty)||0)/1000*r.price;name=r.name;}}else{const pr=prepList.find(pr=>String(pr.id)===String(ing.srcId));if(pr){const {costPerUnit}=calcPrepCost(pr);cost=(parseFloat(ing.qty)||0)/1000*costPerUnit;name=pr.name;}}if(cost>topC2){topC2=cost;topDrv=name;}});
              return <tr key={p.id} style={{borderBottom:`1px solid ${C.border}22`,background:v>0?C.red+"08":"transparent"}}>
                <td style={{padding:"10px 12px",fontWeight:600}}>{p.name}</td>
                <td style={{padding:"10px 12px",color:"#f87171",fontWeight:600}}>{p.totalCost.toFixed(2)}</td>
                <td style={{padding:"10px 12px",color:C.muted}}>{parseFloat(p.stdCost).toFixed(2)}</td>
                <td style={{padding:"10px 12px"}}><span style={{color:v>0?"#f87171":"#4ade80",fontWeight:700}}>{v>0?"+":""}{v.toFixed(2)}</span></td>
                <td style={{padding:"10px 12px"}}><span style={{background:v>0?"#f8717122":"#4ade8022",color:v>0?"#f87171":"#4ade80",padding:"2px 8px",borderRadius:20,fontSize:12,fontWeight:700}}>{vp>0?"+":""}{vp.toFixed(1)}%</span></td>
                <td style={{padding:"10px 12px",color:C.yellow,fontSize:12,fontWeight:600}}>{topDrv}</td>
              </tr>;
            })}</tbody>
          </table>
        </div>
      </div>}

      {/* === CARD 6: Top Cost Drivers === */}
      {show("costanalysis")&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.topCostDriversReport} sub={ar?"أعلى المكونات تكلفة عبر جميع المنتجات":"Top cost ingredients across all products"}/>
        {drivers.length===0?noMsg:<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <div>{drivers.slice(0,Math.ceil(topN/2)).map(([name,cost],i)=><div key={i} style={{marginBottom:9}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:12,fontWeight:600,color:C.text}}>{i+1}. {name}</span><span style={{fontSize:12,color:C.yellow,fontWeight:700}}>{cost.toFixed(2)}</span></div>
            <Bar val={cost} max={maxDrv} color={C.yellow}/>
          </div>)}</div>
          <div>{drivers.slice(Math.ceil(topN/2)).map(([name,cost],i)=><div key={i} style={{marginBottom:9}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:12,fontWeight:600,color:C.text}}>{Math.ceil(topN/2)+i+1}. {name}</span><span style={{fontSize:12,color:C.yellow,fontWeight:700}}>{cost.toFixed(2)}</span></div>
            <Bar val={cost} max={maxDrv} color={C.yellow}/>
          </div>)}</div>
        </div>}
      </div>}

      {/* === CARD 7: Prep Analysis === */}
      {show("prep")&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.secPrep}/>
        {prepCalc.length===0?noMsg:<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <div>
            <div style={{fontSize:11,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:10}}>{t.prepHighCost}</div>
            {[...prepCalc].sort((a,b)=>b.costPerUnit-a.costPerUnit).slice(0,topN).map((p,i)=><div key={p.id} style={{marginBottom:8}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:12,color:C.text,fontWeight:600}}>{i+1}. {p.name}</span><span style={{fontSize:12,color:C.accent,fontWeight:700}}>{p.costPerUnit.toFixed(3)}</span></div>
              <Bar val={p.costPerUnit} max={prepCalc.reduce((a,x)=>Math.max(a,x.costPerUnit),0.001)} color={C.accent}/>
            </div>)}
          </div>
          <div>
            <div style={{fontSize:11,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:10}}>{t.prepMostUsed}</div>
            {[...prepUsage].sort((a,b)=>b.usedIn-a.usedIn).slice(0,topN).map((p,i)=><div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:C.surface,borderRadius:8,marginBottom:6,border:`1px solid ${C.border}`}}>
              <span style={{fontSize:12,fontWeight:600,color:C.text,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{i+1}. {p.name}</span>
              <button onClick={()=>p.usedIn>0&&setRelModal({title:p.name,list:p.relatedProds,type:"products"})} style={{background:p.usedIn>0?C.blue+"22":"transparent",color:p.usedIn>0?C.blue:C.muted,padding:"2px 10px",borderRadius:20,fontSize:12,fontWeight:700,border:"none",cursor:p.usedIn>0?"pointer":"default",fontFamily:"inherit"}}>
                {p.usedIn} {t.productsCount}
              </button>
            </div>)}
          </div>
        </div>}
      </div>}

      {/* === CARD 8: Raw Materials Analysis === */}
      {show("raw")&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.secRaw}/>
        {rawList.length===0?noMsg:<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <div>
            <div style={{fontSize:11,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:10}}>{t.rawHighPrice}</div>
            {[...rawUsage].sort((a,b)=>b.price-a.price).slice(0,topN).map((r,i)=><div key={r.id} style={{marginBottom:8}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:12,color:C.text,fontWeight:600}}>{i+1}. {r.name}</span><span style={{fontSize:12,color:C.accent,fontWeight:700}}>{r.price.toFixed(2)}</span></div>
              <Bar val={r.price} max={rawUsage.reduce((a,x)=>Math.max(a,x.price),0.001)} color={C.accent}/>
            </div>)}
          </div>
          <div>
            <div style={{fontSize:11,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:10}}>{t.rawMostUsed}</div>
            {[...rawUsage].sort((a,b)=>(b.inPrep+b.inProd)-(a.inPrep+a.inProd)).slice(0,topN).map((r,i)=><div key={r.id} style={{padding:"8px 12px",background:C.surface,borderRadius:8,marginBottom:6,border:`1px solid ${C.border}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontSize:12,fontWeight:600,color:C.text,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{i+1}. {r.name}</span>
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>r.inPrep>0&&setRelModal({title:r.name,list:r.relatedPreps,type:"prep"})} style={{background:r.inPrep>0?"#1e3a5f":"transparent",color:r.inPrep>0?"#60a5fa":C.muted,padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700,border:"none",cursor:r.inPrep>0?"pointer":"default",fontFamily:"inherit"}}>P:{r.inPrep}</button>
                  <button onClick={()=>r.inProd>0&&setRelModal({title:r.name,list:r.relatedProds,type:"products"})} style={{background:r.inProd>0?"#0a3326":"transparent",color:r.inProd>0?"#4ade80":C.muted,padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700,border:"none",cursor:r.inProd>0?"pointer":"default",fontFamily:"inherit"}}>M:{r.inProd}</button>
                </div>
              </div>
            </div>)}
          </div>
        </div>}
      </div>}

      {/* RelModal */}
      {relModal&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&setRelModal(null)}>
        <div className="modal" style={{maxWidth:500}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
            <div>
              <div style={{fontWeight:800,fontSize:15,color:C.accent}}>{relModal.title||relModal.cat}</div>
              <div style={{fontSize:12,color:C.muted,marginTop:2}}>{relModal.type==="prep"?t.relatedPreps:t.relatedProducts} ({(relModal.list||relModal.items||[]).length})</div>
            </div>
            <button onClick={()=>setRelModal(null)} style={{background:"transparent",border:"none",color:C.muted,fontSize:20,cursor:"pointer"}}>✕</button>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:8,maxHeight:360,overflowY:"auto"}}>
            {(relModal.list||relModal.items||[]).map((item,i)=><div key={item.id||i} style={{background:C.surface,borderRadius:9,padding:"10px 14px",border:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><code style={{background:C.bg||C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8",marginLeft:8,marginRight:8}}>{item.code}</code><span style={{fontWeight:600,fontSize:13}}>{item.name}</span></div>
              {relModal.type==="classification"&&<span style={{fontSize:12,color:relModal.col,fontWeight:700}}>{getMargin(item).toFixed(1)}%</span>}
            </div>)}
          </div>
          <div style={{display:"flex",justifyContent:"flex-end",marginTop:14}}><button className="btn btn-secondary" onClick={()=>setRelModal(null)}>{t.cancel}</button></div>
        </div>
      </div>}
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════
// SHARED MODALS
// ═══════════════════════════════════════════════════════════════
function DelModal({t,C=DARK,onOk,onCancel}) {
  return <div className="overlay"><div className="modal" style={{maxWidth:300,textAlign:"center"}}>
    <p style={{marginBottom:20,color:C.text,fontSize:14}}>{t.confirmDelete}</p>
    <div style={{display:"flex",gap:10,justifyContent:"center"}}>
      <button className="btn-sm-d" style={{padding:"8px 20px",fontSize:13}} onClick={onOk}>{t.delete}</button>
      <button className="btn btn-secondary" style={{padding:"8px 20px",fontSize:13}} onClick={onCancel}>{t.cancel}</button>
    </div>
  </div></div>;
}

function UsageModal({t,C=DARK,usageModal,onClose}) {
  return <div className="overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{maxWidth:500}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div><div style={{fontWeight:800,fontSize:15,color:C.accent}}>{usageModal.item.name}</div>
          <div style={{fontSize:12,color:C.muted,marginTop:2}}>{usageModal.type==="prep"?t.relatedPreps:t.relatedProducts} ({usageModal.list.length})</div>
        </div>
        <button onClick={onClose} style={{background:"transparent",border:"none",color:C.muted,fontSize:20,cursor:"pointer"}}>✕</button>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:8,maxHeight:360,overflowY:"auto"}}>
        {usageModal.list.map((item,i)=><div key={item.id||i} style={{background:C.surface,borderRadius:9,padding:"10px 14px",border:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8",marginLeft:8,marginRight:8}}>{item.code}</code><span style={{fontWeight:600,fontSize:13}}>{item.name}</span></div>
          <span className="badge badge-cls">{item.class||"—"}</span>
        </div>)}
      </div>
      <div style={{display:"flex",justifyContent:"flex-end",marginTop:14}}><button className="btn btn-secondary" onClick={onClose}>{t.cancel}</button></div>
    </div>
  </div>;
}

// ImportModal with date picker for raw materials
function ImportModal({t,lang,C=DARK,type,onClose,onFileSelect,onDownloadTemplate,showDatePicker,onDateSelect}) {
  const fileRef=useRef();
  const [showDP,setShowDP]=useState(false);
  const [selMonth,setSelMonth]=useState(String(new Date().getMonth()+1).padStart(2,"0"));
  const [selYear,setSelYear]=useState(String(new Date().getFullYear()));
  const ar=lang==="ar";

  const handleFile=(e)=>{
    if(showDatePicker&&onDateSelect){onDateSelect({month:selMonth,year:selYear});}
    onFileSelect(e);
    onClose();
  };

  return <div className="overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{maxWidth:420}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div><div style={{fontWeight:800,fontSize:15,color:C.accent}}>{ar?"استيراد البيانات":"Import Data"}</div><div style={{fontSize:12,color:C.muted,marginTop:2}}>{type}</div></div>
        <button onClick={onClose} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer"}}>✕</button>
      </div>

      {showDatePicker&&<div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,padding:"12px 14px",marginBottom:12}}>
        <div style={{fontSize:12,fontWeight:700,color:C.accent,marginBottom:10}}>{ar?"تحديد تاريخ الأسعار":"Select Price Period"}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
          <div>
            <label style={{fontSize:10,color:C.muted,display:"block",marginBottom:3,fontWeight:700,textTransform:"uppercase"}}>{ar?"الشهر":"Month"}</label>
            <select value={selMonth} onChange={e=>setSelMonth(e.target.value)}>
              {MONTHS_LIST.map(m=><option key={m} value={m}>{mLabel(m,lang)}</option>)}
            </select>
          </div>
          <div>
            <label style={{fontSize:10,color:C.muted,display:"block",marginBottom:3,fontWeight:700,textTransform:"uppercase"}}>{ar?"السنة":"Year"}</label>
            <select value={selYear} onChange={e=>setSelYear(e.target.value)}>
              {["2023","2024","2025","2026","2027"].map(y=><option key={y} value={y}>{y}</option>)}
            </select>
          </div>
        </div>
      </div>}

      <div onClick={onDownloadTemplate} style={{background:C.surface,border:`2px solid ${C.accent}44`,borderRadius:12,padding:"18px 20px",marginBottom:12,cursor:"pointer"}} onMouseOver={e=>{e.currentTarget.style.borderColor=C.accent;}} onMouseOut={e=>{e.currentTarget.style.borderColor=C.accent+"44";}}>
        <div style={{display:"flex",alignItems:"center",gap:14}}><div style={{fontSize:28}}>📥</div><div><div style={{fontWeight:700,color:C.accent,fontSize:13,marginBottom:3}}>{ar?"تحميل النموذج الفارغ":"Download Empty Template"}</div><div style={{fontSize:11,color:C.muted}}>{ar?"احصل على النموذج الصحيح":"Get the correct format"}</div></div></div>
      </div>
      <div onClick={()=>fileRef.current.click()} style={{background:C.surface,border:`2px solid ${C.blue}44`,borderRadius:12,padding:"18px 20px",cursor:"pointer"}} onMouseOver={e=>{e.currentTarget.style.borderColor=C.blue;}} onMouseOut={e=>{e.currentTarget.style.borderColor=C.blue+"44";}}>
        <div style={{display:"flex",alignItems:"center",gap:14}}><div style={{fontSize:28}}>📤</div><div><div style={{fontWeight:700,color:C.blue,fontSize:13,marginBottom:3}}>{ar?"رفع ملف Excel":"Upload Excel File"}</div><div style={{fontSize:11,color:C.muted}}>{ar?"ارفع ملف Excel مملوء":"Upload a filled Excel"}</div></div></div>
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleFile}/>
      <button className="btn btn-secondary" onClick={onClose} style={{width:"100%",marginTop:12,textAlign:"center"}}>{t.cancel}</button>
    </div>
  </div>;
}

function PreviewModal({t,lang,C=DARK,data,onConfirm,onCancel,classes,moduleType}) {
  const [items,setItems]=useState(()=>data.items.map((it,i)=>({...it,_idx:i,_skip:false})));
  const ar=lang==="ar";
  const unitOpts=["kg","liter","piece"];
  const clsOpts=moduleType==="raw"?(classes||[]):moduleType==="prep"?(classes||[]):[];
  const addCount=items.filter(it=>!it._skip&&!it._existing).length;
  const updCount=items.filter(it=>!it._skip&&it._existing).length;
  const skipCount=items.filter(it=>it._skip).length;
  const update=(idx,field,val)=>setItems(prev=>prev.map(it=>it._idx===idx?{...it,[field]:val}:it));
  return <div className="overlay" onClick={e=>e.target===e.currentTarget&&onCancel()}>
    <div className="modal" style={{maxWidth:900,width:"100%",maxHeight:"90vh",display:"flex",flexDirection:"column"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexShrink:0}}>
        <div>
          <div style={{fontWeight:800,fontSize:15,color:C.accent}}>{ar?"مراجعة البيانات":"Review Before Import"}</div>
          <div style={{fontSize:12,color:C.muted,marginTop:3,display:"flex",gap:12}}>
            <span style={{color:C.green}}>✓ {ar?"يضاف":"Add"}: {addCount}</span>
            <span style={{color:C.yellow}}>↻ {ar?"يحدَّث":"Update"}: {updCount}</span>
            <span style={{color:C.muted}}>⊘ {ar?"يتجاهل":"Skip"}: {skipCount}</span>
          </div>
        </div>
        <button onClick={onCancel} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer"}}>✕</button>
      </div>
      <div style={{overflowY:"auto",flex:1,marginBottom:14}}>
        <table style={{borderCollapse:"collapse",fontSize:12}}>
          <thead style={{position:"sticky",top:0,zIndex:10}}>
            <tr>{["#",ar?"الاسم":"Name",ar?"الكلاس":"Class",ar?"الوحدة":"Unit",ar?"السعر":"Price",ar?"الإجراء":"Action",ar?"تجاهل":"Skip"].map((h,i)=>(
              <th key={i} style={{padding:"8px 10px",textAlign:ar?"right":"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",background:C.surface,borderBottom:`1px solid ${C.border}`}}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {items.map((it)=>(
              <tr key={it._idx} style={{opacity:it._skip?0.4:1,background:it._skip?"transparent":it._existing?C.accent+"08":C.green+"08"}}>
                <td style={{padding:"7px 10px",color:C.muted,fontSize:11}}>{it._idx+1}</td>
                <td style={{padding:"7px 10px"}}><input value={it.name} onChange={e=>update(it._idx,"name",e.target.value)} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"4px 8px",fontSize:12,width:"100%",outline:"none",minWidth:120}}/></td>
                <td style={{padding:"7px 10px"}}>
                  {clsOpts.length>0
                    ?<select value={it.class||""} onChange={e=>update(it._idx,"class",e.target.value)} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"4px 8px",fontSize:12,width:"100%",outline:"none"}}>
                      <option value="">—</option>{clsOpts.map(c=><option key={c} value={c}>{c}</option>)}
                    </select>
                    :<span style={{color:C.muted,fontSize:11}}>—</span>}
                </td>
                <td style={{padding:"7px 10px"}}>
                  <select value={it.unit||"kg"} onChange={e=>update(it._idx,"unit",e.target.value)} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"4px 8px",fontSize:12,width:"100%",outline:"none"}}>
                    {unitOpts.map(u=><option key={u} value={u}>{u==="kg"?t.kg:u==="liter"?t.liter:t.piece}</option>)}
                  </select>
                </td>
                <td style={{padding:"7px 10px"}}>
                  {it.price!==undefined
                    ?<input type="number" value={it.price||0} onChange={e=>update(it._idx,"price",e.target.value)} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.accent,borderRadius:5,padding:"4px 8px",fontSize:12,width:80,outline:"none",fontWeight:700}}/>
                    :<span style={{color:C.muted}}>—</span>}
                </td>
                <td style={{padding:"7px 10px"}}>
                  <span style={{padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700,background:it._existing?C.yellow+"22":C.green+"22",color:it._existing?C.yellow:C.green}}>
                    {it._existing?(ar?"تحديث":"Update"):(ar?"إضافة":"Add")}
                  </span>
                </td>
                <td style={{padding:"7px 10px",textAlign:"center"}}>
                  <input type="checkbox" checked={!!it._skip} onChange={e=>update(it._idx,"_skip",e.target.checked)} style={{width:15,height:15,accentColor:C.danger,cursor:"pointer"}}/>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{display:"flex",gap:8,flexShrink:0,borderTop:`1px solid ${C.border}`,paddingTop:12}}>
        <button className="btn btn-primary" style={{flex:2}} onClick={()=>onConfirm(items)}>
          {ar?`✅ تأكيد الاستيراد (${addCount+updCount} صنف)`:`✅ Confirm Import (${addCount+updCount} items)`}
        </button>
        <button className="btn btn-secondary" style={{flex:1}} onClick={onCancel}>{t.cancel}</button>
      </div>
    </div>
  </div>;
}

function DupDeleteModal({t,lang,C=DARK,items,onDelete,onClose}) {
  const ar=lang==="ar";
  const groups={};
  items.forEach(it=>{const k=it.name.trim().toLowerCase();if(!groups[k])groups[k]=[];groups[k].push(it);});
  const [toDelete,setToDelete]=useState({});
  const toggleDel=(id)=>setToDelete(p=>({...p,[id]:!p[id]}));
  const selectedCount=Object.values(toDelete).filter(Boolean).length;
  return <div className="overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{maxWidth:680,width:"100%",maxHeight:"85vh",display:"flex",flexDirection:"column"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexShrink:0}}>
        <div><div style={{fontWeight:800,fontSize:15,color:C.yellow}}>{ar?"إدارة الأسماء المكررة":"Manage Duplicates"}</div></div>
        <button onClick={onClose} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer"}}>✕</button>
      </div>
      <div style={{overflowY:"auto",flex:1,marginBottom:14}}>
        {Object.entries(groups).map(([key,grp])=>(
          <div key={key} style={{marginBottom:14,background:C.surface,borderRadius:10,padding:12,border:`1px solid ${C.yellow}44`}}>
            <div style={{fontSize:12,fontWeight:700,color:C.yellow,marginBottom:8}}>⚠ {ar?"اسم مكرر":"Duplicate"}: <span style={{color:C.text}}>{grp[0].name}</span></div>
            {grp.map((it,i)=>(
              <div key={it.id} style={{display:"flex",alignItems:"center",gap:10,padding:"7px 10px",borderRadius:7,background:toDelete[it.id]?C.danger+"18":C.card,border:`1px solid ${toDelete[it.id]?C.danger+"66":C.border}`,marginBottom:5}}>
                <input type="checkbox" checked={!!toDelete[it.id]} onChange={()=>toggleDel(it.id)} style={{width:14,height:14,accentColor:C.danger,cursor:"pointer"}}/>
                <code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{it.code}</code>
                <span style={{fontWeight:600,fontSize:12,flex:1}}>{it.name}</span>
                <span className="badge badge-cls">{it.class||"—"}</span>
                <span style={{fontSize:12,color:C.accent,fontWeight:700}}>{it.price?.toFixed(2)}</span>
                {i===0&&<span style={{fontSize:10,color:C.green,fontWeight:700,padding:"1px 7px",borderRadius:20,background:C.green+"22"}}>{ar?"الأحدث":"Newest"}</span>}
              </div>
            ))}
          </div>
        ))}
      </div>
      <div style={{display:"flex",gap:8,flexShrink:0,borderTop:`1px solid ${C.border}`,paddingTop:12}}>
        <button className="btn-sm-d" style={{flex:2,padding:"9px",fontSize:13,borderRadius:7,opacity:selectedCount===0?0.5:1,cursor:selectedCount===0?"not-allowed":"pointer"}}
          onClick={()=>{if(selectedCount>0)onDelete(Object.keys(toDelete).filter(id=>toDelete[id]));}}>
          {ar?`حذف ${selectedCount} سجل`:`Delete ${selectedCount}`}
        </button>
        <button className="btn btn-secondary" style={{flex:1}} onClick={onClose}>{t.cancel}</button>
      </div>
    </div>
  </div>;
}


// ═══════════════════════════════════════════════════════════════
// INGREDIENT ROW — Fixed: strong search filter, no unit-changes-material bug
// ═══════════════════════════════════════════════════════════════
function IngRow({ing,rawList,prepList,lang,t,C=DARK,onUpdate,onRemove}) {
  const [q,setQ]=useState("");
  const [open,setOpen]=useState(false);
  const dropRef=useRef(null);
  const inputRef=useRef(null);
  const srcList=ing.source==="prep"?(prepList||[]):(rawList||[]);

  // FIX: Strong search — items starting with query come first, then contains
  const filtered=useMemo(()=>{
    if(!q.trim()) return srcList;
    const lq=q.toLowerCase();
    const starts=srcList.filter(r=>r.name.toLowerCase().startsWith(lq)||r.code?.toLowerCase().startsWith(lq));
    const contains=srcList.filter(r=>!r.name.toLowerCase().startsWith(lq)&&!r.code?.toLowerCase().startsWith(lq)&&(r.name.toLowerCase().includes(lq)||r.code?.toLowerCase().includes(lq)));
    return [...starts,...contains];
  },[srcList,q]);

  const selected=srcList.find(r=>String(r.id)===String(ing.rawId));
  useEffect(()=>{setQ(""); setOpen(false);},[ing.source]);
  useEffect(()=>{
    const close=(e)=>{ if(dropRef.current&&!dropRef.current.contains(e.target))setOpen(false); };
    document.addEventListener("mousedown",close);
    return ()=>document.removeEventListener("mousedown",close);
  },[]);

  const baseUnit=selected?.unit||"kg";
  const inputUnitOpts=INPUT_UNITS[baseUnit]||INPUT_UNITS.kg;
  const currentInputUnit=ing.inputUnit||inputUnitOpts[0].val;

  // FIX: Changing unit does NOT change material — separate handler
  const handleUnitChange=(iu)=>{ onUpdate("inputUnit",iu); };
  const handleMaterialSelect=(id)=>{ 
    onUpdate("rawId",id); 
    onUpdate("inputUnit","");
    setOpen(false); 
    setQ(""); 
  };

  return (
    <div ref={dropRef} style={{display:"grid",gap:7,gridTemplateColumns:"1fr 2fr 1fr 1fr 1fr auto",alignItems:"end",marginBottom:7,background:C.surface,padding:9,borderRadius:8,border:`1px solid ${C.border}`}}>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.source}</label>
        <select value={ing.source||"raw"} onChange={e=>{onUpdate("source",e.target.value);onUpdate("rawId","");onUpdate("inputUnit","");}} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}>
          <option value="raw">{lang==="ar"?"مادة خام":"Raw"}</option>
          <option value="prep">{lang==="ar"?"بريب":"Prep"}</option>
        </select>
      </div>
      <div style={{position:"relative"}}>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.ingredient}</label>
        <div onClick={()=>setOpen(o=>!o)} style={{background:C.surface,border:`1px solid ${open?C.accent:C.border}`,borderRadius:7,padding:"8px 10px",cursor:"pointer",fontSize:12,color:selected?C.text:C.muted,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{selected?selected.name+" ("+selected.code+")":"—"}</span>
          <span style={{fontSize:9}}>▼</span>
        </div>
        {open&&<div style={{position:"absolute",top:"100%",left:0,right:0,zIndex:200,background:C.card,border:`1px solid ${C.border}`,borderRadius:8,boxShadow:"0 8px 32px rgba(0,0,0,.6)",marginTop:2}}>
          <div style={{padding:6}}><input ref={inputRef} autoFocus placeholder={lang==="ar"?"بحث...":"Search..."} value={q} onChange={e=>setQ(e.target.value)} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"6px 10px",fontSize:11,outline:"none",width:"100%"}}/></div>
          <div style={{maxHeight:200,overflowY:"auto"}}>
            <div onMouseDown={(e)=>{e.preventDefault();e.stopPropagation();handleMaterialSelect("");}} style={{padding:"7px 10px",cursor:"pointer",fontSize:12,color:C.muted}}>—</div>
            {filtered.map(r=><div key={r.id} onMouseDown={(e)=>{e.preventDefault();e.stopPropagation();handleMaterialSelect(r.id);}} style={{padding:"7px 10px",cursor:"pointer",fontSize:12,color:String(r.id)===String(ing.rawId)?C.accent:C.text,background:String(r.id)===String(ing.rawId)?C.accent+"15":"transparent"}}>{r.name} <span style={{color:C.muted,fontSize:10}}>({r.code})</span> <span style={{color:C.blue,fontSize:10}}>[{r.unit}]</span></div>)}
            {filtered.length===0&&<div style={{padding:"7px 10px",color:C.muted,fontSize:11}}>{t.noData}</div>}
          </div>
        </div>}
        {selected&&<div style={{fontSize:10,color:C.muted,marginTop:2}}>{lang==="ar"?"وحدة الأساس:":"Base unit:"} <span style={{color:C.accent}}>{selected.unit}</span></div>}
      </div>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{lang==="ar"?"الوحدة":"Unit"}</label>
        <select value={currentInputUnit} onChange={e=>handleUnitChange(e.target.value)} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}>
          {inputUnitOpts.map(u=><option key={u.val} value={u.val}>{u.label}</option>)}
        </select>
      </div>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.qty}</label>
        <input type="number" min="0" step="0.001" value={ing.qty||""} onChange={e=>onUpdate("qty",e.target.value)} placeholder="0" style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}/>
        {selected&&parseFloat(ing.qty)>0&&(()=>{
          const internal=toInternal(parseFloat(ing.qty)||0,currentInputUnit);
          return <div style={{fontSize:10,color:C.muted,marginTop:2}}>= {internal.toFixed(1)} {baseUnit==="kg"?"g":baseUnit==="liter"?"ml":"pcs"}</div>;
        })()}
      </div>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.waste} %</label>
        <input type="number" min="0" max="100" step="0.5" value={ing.waste||""} onChange={e=>onUpdate("waste",e.target.value)} placeholder="0" style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}/>
      </div>
      <button onClick={onRemove} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"4px 10px",cursor:"pointer",marginTop:22,fontFamily:"inherit",fontWeight:600,fontSize:12}}>✕</button>
    </div>
  );
}

// ─── INGREDIENT ROW FOR PRODUCTS ─────────────────────────────
function IngRowProd({ing,rawList,prepList,lang,t,C=DARK,onUpdate,onRemove}) {
  const [q,setQ]=useState(""); const [open,setOpen]=useState(false);
  const dropRef=useRef(null);
  const srcList=ing.source==="raw"?rawList:prepList;

  // FIX: Strong search — exact starts-with first
  const filtered=useMemo(()=>{
    if(!q.trim()) return srcList;
    const lq=q.toLowerCase();
    const starts=srcList.filter(r=>r.name.toLowerCase().startsWith(lq)||r.code?.toLowerCase().startsWith(lq));
    const contains=srcList.filter(r=>!r.name.toLowerCase().startsWith(lq)&&!r.code?.toLowerCase().startsWith(lq)&&(r.name.toLowerCase().includes(lq)||r.code?.toLowerCase().includes(lq)));
    return [...starts,...contains];
  },[srcList,q]);

  const selected=srcList.find(r=>String(r.id)===String(ing.srcId));
  useEffect(()=>{setQ(""); setOpen(false);},[ing.source]);
  useEffect(()=>{
    const close=(e)=>{ if(dropRef.current&&!dropRef.current.contains(e.target))setOpen(false); };
    document.addEventListener("mousedown",close);
    return ()=>document.removeEventListener("mousedown",close);
  },[]);
  const baseUnit=selected?.unit||"kg";
  const inputUnitOpts=INPUT_UNITS[baseUnit]||INPUT_UNITS.kg;
  const currentInputUnit=ing.inputUnit||inputUnitOpts[0].val;

  // FIX: Changing unit does NOT change material
  const handleUnitChange=(iu)=>{ onUpdate("inputUnit",iu); };
  const handleMaterialSelect=(id)=>{
    onUpdate("srcId", id ? String(id) : "");
    onUpdate("inputUnit","");
    setOpen(false);
    setQ("");
  };

  return (
    <div ref={dropRef} style={{display:"grid",gap:7,gridTemplateColumns:"1fr 2fr 1fr 1fr 1fr auto",alignItems:"end",marginBottom:7,background:C.surface,padding:9,borderRadius:8,border:`1px solid ${C.border}`}}>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.source}</label>
        <select value={ing.source||"raw"} onChange={e=>{onUpdate("source",e.target.value);onUpdate("srcId","");onUpdate("inputUnit","");}} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}>
          <option value="raw">{t.rawMat}</option>
          <option value="prep">{t.prepItem}</option>
        </select>
      </div>
      <div style={{position:"relative"}}>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.ingredient}</label>
        <div onClick={()=>setOpen(o=>!o)} style={{background:C.surface,border:`1px solid ${open?C.accent:C.border}`,borderRadius:7,padding:"8px 10px",cursor:"pointer",fontSize:12,color:selected?C.text:C.muted,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{selected?selected.name+" ("+selected.code+")":"—"}</span>
          <span style={{fontSize:9}}>▼</span>
        </div>
        {open&&<div style={{position:"absolute",top:"100%",left:0,right:0,zIndex:200,background:C.card,border:`1px solid ${C.border}`,borderRadius:8,boxShadow:"0 8px 32px rgba(0,0,0,.6)",marginTop:2}}>
          <div style={{padding:6}}><input autoFocus placeholder={lang==="ar"?"بحث...":"Search..."} value={q} onChange={e=>setQ(e.target.value)} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"6px 10px",fontSize:11,outline:"none",width:"100%"}}/></div>
          <div style={{maxHeight:200,overflowY:"auto"}}>
            <div onMouseDown={(e)=>{e.preventDefault();e.stopPropagation();handleMaterialSelect("");}} style={{padding:"7px 10px",cursor:"pointer",fontSize:12,color:C.muted}}>—</div>
            {filtered.map(r=><div key={r.id} onMouseDown={(e)=>{e.preventDefault();e.stopPropagation();handleMaterialSelect(r.id);}} style={{padding:"7px 10px",cursor:"pointer",fontSize:12,color:String(r.id)===String(ing.srcId)?C.accent:C.text,background:String(r.id)===String(ing.srcId)?C.accent+"15":"transparent"}}>{r.name} <span style={{color:C.muted,fontSize:10}}>({r.code})</span> <span style={{color:C.blue,fontSize:10}}>[{r.unit}]</span></div>)}
            {filtered.length===0&&<div style={{padding:"7px 10px",color:C.muted,fontSize:11}}>{t.noData}</div>}
          </div>
        </div>}
        {selected&&<div style={{fontSize:10,color:C.muted,marginTop:2}}>{lang==="ar"?"وحدة:":"Unit:"} <span style={{color:C.accent}}>{selected.unit}</span></div>}
      </div>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{lang==="ar"?"الوحدة":"Unit"}</label>
        <select value={currentInputUnit} onChange={e=>handleUnitChange(e.target.value)} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}>
          {inputUnitOpts.map(u=><option key={u.val} value={u.val}>{u.label}</option>)}
        </select>
      </div>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.qty}</label>
        <input type="number" min="0" step="0.001" value={ing.qty||""} onChange={e=>onUpdate("qty",e.target.value)} placeholder="0" style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}/>
        {selected&&parseFloat(ing.qty)>0&&(()=>{
          const internal=toInternal(parseFloat(ing.qty)||0,currentInputUnit);
          return <div style={{fontSize:10,color:C.muted,marginTop:2}}>= {internal.toFixed(1)} {baseUnit==="kg"?"g":baseUnit==="liter"?"ml":"pcs"}</div>;
        })()}
      </div>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.waste} %</label>
        <input type="number" min="0" max="100" step="0.5" value={ing.waste||""} onChange={e=>onUpdate("waste",e.target.value)} placeholder="0" style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}/>
      </div>
      <button onClick={onRemove} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"4px 10px",cursor:"pointer",marginTop:22,fontFamily:"inherit",fontWeight:600,fontSize:12}}>✕</button>
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════
// RAW MATERIALS TAB — with date picker on import
// ═══════════════════════════════════════════════════════════════
function RawTab({t,lang,C=DARK,rawList,setRawList,classes,prepList=[],prodList=[],showToast,hasPerm,mod}) {
  const [searchRaw,setSearchRaw]=useState(""); const [search,setSearch]=useState("");
  const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",unit:"kg",price:"",class:""});
  const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null);
  const [usageModal,setUsageModal]=useState(null);
  const [showImport,setShowImport]=useState(false);
  const [previewData,setPreviewData]=useState(null);
  const [bulkDelModal,setBulkDelModal]=useState(null);
  const [pageSize,setPageSize]=useState(20); const [showCount,setShowCount]=useState(20);
  const [importDate,setImportDate]=useState(null);
  const fileRef=useRef();
  const cls=classes.raw||[];

  useEffect(()=>{const timer=setTimeout(()=>setSearch(searchRaw),200);return()=>clearTimeout(timer);},[searchRaw]);
  useEffect(()=>setShowCount(pageSize),[search,fcls,pageSize]);

  const ok=()=>{
    const e={};
    if(!form.name.trim()) e.name=t.required;
    if(!form.price||parseFloat(form.price)<=0) e.price=t.positiveNum;
    const nameExists=rawList.some(m=>m.name.trim().toLowerCase()===form.name.trim().toLowerCase()&&m.id!==editId);
    if(nameExists) e.name=lang==="ar"?"هذا الاسم موجود مسبقاً":"Name already exists";
    setErrs(e); return !Object.keys(e).length;
  };
  const reset=()=>{ setForm({name:"",unit:"kg",price:"",class:""}); setErrs({}); setShowForm(false); setEditId(null); };
  const save=()=>{
    if(!ok()) return;
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    if(editId!==null) setRawList(p=>p.map(m=>m.id===editId?{...m,name:form.name.trim(),unit:form.unit,price:parseFloat(form.price),class:form.class,lastUpdated:now}:m));
    else {
      const code=genCode("Raw",rawList);
      setRawList(p=>[...p,{id:Date.now(),code,name:form.name.trim(),unit:form.unit,price:parseFloat(form.price),class:form.class,lastUpdated:now}]);
    }
    reset(); showToast(t.savedOk);
  };
  const doEdit=m=>{ setForm({name:m.name,unit:m.unit,price:String(m.price),class:m.class||""}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setRawList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };

  const findDuplicates=()=>{
    const nameMap={};
    rawList.forEach(m=>{const k=m.name.trim().toLowerCase();if(!nameMap[k])nameMap[k]=[];nameMap[k].push(m);});
    return Object.values(nameMap).filter(g=>g.length>1).flat();
  };
  const dups=findDuplicates();

  const doDownloadTemplate=()=>{
    const sample=[
      {Code:"Raw-00001",Name:lang==="ar"?"دجاج خام":"Chicken Raw",Class:"Food Item",Unit:"kg",Price:20},
      {Code:"Raw-00002",Name:lang==="ar"?"طماطم":"Tomato",Class:"Food Item",Unit:"kg",Price:5},
      {Code:"Raw-00003",Name:lang==="ar"?"زيت نباتي":"Vegetable Oil",Class:"Food Item",Unit:"liter",Price:8},
    ];
    const ws=XLSX.utils.json_to_sheet(sample); ws["!cols"]=[{wch:14},{wch:30},{wch:14},{wch:10},{wch:12}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Raw Materials");
    XLSX.writeFile(wb,"TALA_Raw_Materials_Template.xlsx");
  };
  const doExport=()=>{
    const data=rawList.map(m=>({Code:m.code,Name:m.name,Class:m.class||"",Unit:m.unit,Price:m.price}));
    if(!data.length){doDownloadTemplate();return;}
    const ws=XLSX.utils.json_to_sheet(data); ws["!cols"]=[{wch:14},{wch:30},{wch:14},{wch:10},{wch:12}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Raw Materials");
    XLSX.writeFile(wb,"TALA_Raw_Materials_Export.xlsx");
  };
  const doImport=e=>{
    const file=e.target.files[0]; if(!file) return;
    const r=new FileReader();
    r.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        if(!rows.length){showToast(t.noMatch,"warning");return;}
        const fv=(row,keys)=>{ for(const k of keys){const v=row[k];if(v!==undefined&&v!==null&&String(v).trim()!=="")return String(v).trim();}return ""; };
        const items=[];
        rows.forEach(row=>{
          const rn=fv(row,["Name","الاسم","name","اسم الصنف","Item Name","المادة","المادة الخام","Raw Material"]);
          const rc=fv(row,["Code","الكود","code","كود الصنف","Item Code","الرمز"]);
          const rpRaw=fv(row,["Price","السعر","price","Cost","التكلفة","تكلفة","سعر","Cost Per KG","New Cost","pricePerKg"]);
          const rp=parseFloat(rpRaw)||0;
          const ruRaw=fv(row,["Unit","الوحدة","unit","وحدة","UOM"]);
          const ru=ruRaw.toLowerCase();
          const rcls=fv(row,["Class","الكلاس","class","كلاس","Category","فئة","Type","النوع"])||"Food Item";
          if(!rn) return;
          const unit=["liter","litre","l","لتر"].includes(ru)?"liter":["piece","pieces","pcs","حبة","قطعة"].includes(ru)?"piece":"kg";
          const existing=rawList.find(m=>m.name.toLowerCase()===rn.toLowerCase()) || (rc?rawList.find(m=>m.code===rc):null);
          items.push({name:rn,code:rc||"",unit,price:rp,class:rcls,_existing:existing});
        });
        setPreviewData({items,module:"raw",importDate});
        setShowImport(false);
      }catch(err){showToast(lang==="ar"?"خطأ في الملف":"File error","error");}
    };
    r.readAsBinaryString(file); e.target.value="";
  };
  const confirmImport=(editedItems)=>{
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    const dateLabel=importDate?`${mLabel(importDate.month,lang)} ${importDate.year}`:now;
    let updated=0,added=0;
    setRawList(prev=>{
      const u=[...prev];
      editedItems.forEach(item=>{
        if(item._skip) return;
        const i=item._existing?u.findIndex(m=>m.id===item._existing.id):-1;
        if(i!==-1){
          u[i]={...u[i],name:item.name.trim(),price:parseFloat(item.price)||0,unit:item.unit,class:item.class||u[i].class,lastUpdated:dateLabel};
          updated++;
        } else {
          const nameExists=u.some(m=>m.name.trim().toLowerCase()===item.name.trim().toLowerCase());
          if(nameExists) return;
          const newCode=item.code&&!["undefined","nan",""].includes(item.code)?item.code:genCode("Raw",u);
          u.push({id:Date.now()+Math.random(),code:newCode,name:item.name.trim(),unit:item.unit,price:parseFloat(item.price)||0,class:item.class||"Food Item",lastUpdated:dateLabel});
          added++;
        }
      });
      return u;
    });
    setPreviewData(null);
    setImportDate(null);
    showToast(lang==="ar"?`✅ أضيف ${added} | حُدِّث ${updated}`:`✅ Added ${added} | Updated ${updated}`,(added+updated)>0?"success":"warning");
  };

  const [selected,setSelected]=useState(new Set());
  const toggleSel=(id)=>setSelected(p=>{const n=new Set(p);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>setSelected(p=>p.size===displayed.length?new Set():new Set(displayed.map(m=>m.id)));
  const bulkDelete=()=>{setRawList(p=>p.filter(m=>!selected.has(m.id)));setSelected(new Set());showToast(lang==="ar"?`تم حذف ${selected.size} صنف`:`Deleted ${selected.size} items`,"error");};

  // FIX: class filter only shows exact match
  const filtered=rawList.filter(m=>(m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase()))&&(fcls==="all"||m.class===fcls));
  const displayed=filtered.slice(0,showCount);
  const hasMore=filtered.length>showCount;

  return (
    <div>
      {dups.length>0&&<div style={{background:"#3d2205",border:"1px solid #ca8a04",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:16}}>⚠️</span>
          <span style={{fontSize:12,color:"#fcd34d",fontWeight:700}}>{lang==="ar"?`يوجد ${dups.length} مادة خام بأسماء مكررة`:`${dups.length} duplicate raw material names found`}</span>
        </div>
        <button onClick={()=>setBulkDelModal({items:dups,module:"raw"})} style={{background:"#ca8a04",color:"#060810",border:"none",borderRadius:6,padding:"5px 12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{lang==="ar"?"عرض وإدارة التكرارات":"View & Manage Duplicates"}</button>
      </div>}
      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:240}} placeholder={t.search} value={searchRaw} onChange={e=>{setSearchRaw(e.target.value);setSelected(new Set());}}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{display:"flex",alignItems:"center",gap:5,background:C.surface,borderRadius:7,padding:"4px 8px",border:`1px solid ${C.border}`}}>
            <span style={{fontSize:11,color:C.muted}}>{t.pageSize}:</span>
            {[20,50,100,200,300].map(n=><button key={n} onClick={()=>{setPageSize(n);setShowCount(n);}} style={{background:pageSize===n?C.accent:"transparent",color:pageSize===n?"#080b14":C.muted,border:"none",borderRadius:5,padding:"3px 8px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{n}</button>)}
          </div>
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={()=>setShowImport(true)}>{t.importXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>
      {selected.size>0&&hasPerm(mod,"delete")&&<div style={{background:"#1a0f0f",border:"1px solid #dc262666",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <span style={{fontSize:13,color:"#fca5a5",fontWeight:700}}>{lang==="ar"?`تم تحديد ${selected.size} صنف`:`${selected.size} selected`}</span>
        <div style={{display:"flex",gap:7}}>
          <button onClick={()=>setSelected(new Set())} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:6,padding:"5px 12px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{lang==="ar"?"إلغاء":"Deselect"}</button>
          <button onClick={bulkDelete} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"5px 14px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>{lang==="ar"?`حذف ${selected.size}`:`Delete ${selected.size}`}</button>
        </div>
      </div>}
      <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <button className={`fbtn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all} ({rawList.length})</button>
        {cls.map(c=><button key={c} className={`fbtn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c} ({rawList.filter(m=>m.class===c).length})</button>)}
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{borderCollapse:"collapse"}}>
          <thead><tr>
            {hasPerm(mod,"delete")&&<th style={{width:36,padding:"10px 8px"}}><input type="checkbox" checked={displayed.length>0&&selected.size===displayed.length} onChange={toggleAll} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></th>}
            {["#",t.code,t.name,t.class,t.unit,t.price,t.usedInPrep,t.usedInProducts,(hasPerm(mod,"edit")||hasPerm(mod,"delete"))?t.actions:""].filter(Boolean).map((h,i)=><th key={i}>{h}</th>)}
          </tr></thead>
          <tbody>
            {displayed.length===0?<tr><td colSpan={10} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :displayed.map((m,i)=>{
              const nPrep=prepList.filter(p=>p.ingredients?.some(i=>String(i.rawId)===String(m.id))).length;
              const nProd=prodList.filter(p=>p.ingredients?.some(i=>i.source==="raw"&&String(i.srcId)===String(m.id))).length;
              const isDup=dups.some(d=>d.id===m.id);
              const isSel=selected.has(m.id);
              return <tr key={m.id} style={{background:isSel?C.accent+"12":isDup?"#3d220511":""}}>
                {hasPerm(mod,"delete")&&<td style={{padding:"10px 8px",textAlign:"center"}}><input type="checkbox" checked={isSel} onChange={()=>toggleSel(m.id)} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></td>}
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>{isDup&&<span style={{color:"#f59e0b",marginLeft:4,marginRight:4,fontSize:12}}>⚠</span>}{m.name}</td>
                <td><span className="badge badge-cls">{m.class||"—"}</span></td>
                <td><span className={`badge badge-${m.unit}`}>{unitLbl(m.unit,t)}</span></td>
                <td style={{color:C.accent,fontWeight:700}}>{m.price.toFixed(2)}</td>
                <td>{nPrep>0?<button onClick={()=>setUsageModal({item:m,type:"prep",list:prepList.filter(p=>p.ingredients?.some(i=>String(i.rawId)===String(m.id)))})} style={{background:"#0f2a4a",color:"#60a5fa",border:"1px solid #1e3a6033",borderRadius:6,padding:"3px 10px",cursor:"pointer",fontWeight:700,fontSize:13}}>{nPrep}</button>:<span style={{color:C.muted}}>0</span>}</td>
                <td>{nProd>0?<button onClick={()=>setUsageModal({item:m,type:"product",list:prodList.filter(p=>p.ingredients?.some(i=>i.source==="raw"&&String(i.srcId)===String(m.id)))})} style={{background:"#0a2a1a",color:"#4ade80",border:"1px solid #16a34a33",borderRadius:6,padding:"3px 10px",cursor:"pointer",fontWeight:700,fontSize:13}}>{nProd}</button>:<span style={{color:C.muted}}>0</span>}</td>
                {(hasPerm(mod,"edit")||hasPerm(mod,"delete"))&&<td><div style={{display:"flex",gap:5}}>{hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>doEdit(m)}>{t.edit}</button>}{hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>setDelId(m.id)}>{t.delete}</button>}</div></td>}
              </tr>;
            })}
          </tbody>
        </table>
      </div></div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginTop:10}}>
        <span style={{fontSize:12,color:C.muted}}>{lang==="ar"?`عرض ${displayed.length} من ${filtered.length}`:`Showing ${displayed.length} of ${filtered.length}`}</span>
        {hasMore&&<button onClick={()=>setShowCount(c=>c+pageSize)} style={{background:C.surface,color:C.accent,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{t.showMore} (+{Math.min(pageSize,filtered.length-showCount)})</button>}
        {showCount>pageSize&&!hasMore&&<button onClick={()=>setShowCount(pageSize)} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{t.showLess}</button>}
      </div>
      {showImport&&<ImportModal t={t} lang={lang} C={C} type={t.rawMat} onClose={()=>setShowImport(false)} onDownloadTemplate={doDownloadTemplate} onFileSelect={doImport} showDatePicker={true} onDateSelect={setImportDate}/>}
      {usageModal&&<UsageModal t={t} C={C} usageModal={usageModal} onClose={()=>setUsageModal(null)}/>}
      {previewData&&previewData.module==="raw"&&<PreviewModal t={t} lang={lang} C={C} data={previewData} onConfirm={confirmImport} onCancel={()=>setPreviewData(null)} classes={cls} moduleType="raw"/>}
      {bulkDelModal&&bulkDelModal.module==="raw"&&<DupDeleteModal t={t} lang={lang} C={C} items={bulkDelModal.items} onDelete={ids=>{setRawList(p=>p.filter(m=>!ids.includes(m.id)));setBulkDelModal(null);showToast(lang==="ar"?`تم حذف ${ids.length} مادة مكررة`:`Deleted ${ids.length} duplicates`,"error");}} onClose={()=>setBulkDelModal(null)}/>}
      {showForm&&<div className="overlay" onClick={e=>{if(e.target===e.currentTarget)reset();}}><div className="modal" onClick={e=>e.stopPropagation()}>
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.edit:t.add} — {t.rawMat}</h2>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <div><label className="lbl">{t.name}</label><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">{t.class}</label><select value={form.class} onChange={e=>setForm({...form,class:e.target.value})}><option value="">—</option>{cls.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div><label className="lbl">{t.unit}</label><select value={form.unit} onChange={e=>setForm({...form,unit:e.target.value})}><option value="kg">{t.kg}</option><option value="liter">{t.liter}</option><option value="piece">{t.piece}</option></select></div>
            <div><label className="lbl">{t.price} / {t.unit}</label><input type="number" min="0" step="0.01" value={form.price} onChange={e=>setForm({...form,price:e.target.value})} placeholder="0.00"/>{errs.price&&<div className="err">{errs.price}</div>}</div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:4}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
        </div>
      </div></div>}
      {delId&&<DelModal t={t} C={C} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════
// PREP TAB — Fixed: class filter, search filter, ingredient issues
// ═══════════════════════════════════════════════════════════════
function PrepTab({t,lang,C=DARK,prepList,setPrepList,rawList,prodList=[],classes,calcPrepCost,showToast,hasPerm,mod}) {
  const [search,setSearch]=useState(""); const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",unit:"kg",class:"",yieldOverride:"",ingredients:[]}); const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null); const [viewItem,setViewItem]=useState(null);
  const [usageModal,setUsageModal]=useState(null); const [showImport,setShowImport]=useState(false);
  const [previewData,setPreviewData]=useState(null); const [bulkDelModal,setBulkDelModal]=useState(null);
  const [selected,setSelected]=useState(new Set());
  const [pageSize,setPageSize]=useState(20); const [showCount,setShowCount]=useState(20);
  const fileRef=useRef(); const cls=classes.prep||[];

  const findDuplicates=()=>{ const m={}; prepList.forEach(x=>{const k=x.name.trim().toLowerCase();if(!m[k])m[k]=[];m[k].push(x);}); return Object.values(m).filter(g=>g.length>1).flat(); };
  const dups=findDuplicates();
  const blank=()=>({id:Date.now()+Math.random(),source:"raw",rawId:"",qty:"",waste:"0",inputUnit:""});
  const reset=()=>{ setForm({name:"",unit:"kg",class:"",yieldOverride:"",ingredients:[]}); setErrs({}); setShowForm(false); setEditId(null); };
  const ok=()=>{
    const e={};
    if(!form.name.trim()) e.name=t.required;
    const nameExists=prepList.some(m=>m.name.trim().toLowerCase()===form.name.trim().toLowerCase()&&m.id!==editId);
    if(nameExists) e.name=lang==="ar"?"هذا الاسم موجود مسبقاً":"Name already exists";
    setErrs(e); return !Object.keys(e).length;
  };

  const save=()=>{
    if(!ok()) return;
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    const ings=form.ingredients.filter(i=>i.rawId&&parseFloat(i.qty)>0);
    if(editId!==null) setPrepList(p=>p.map(m=>m.id===editId?{...m,name:form.name.trim(),unit:form.unit,class:form.class,yieldOverride:form.yieldOverride,ingredients:ings,lastUpdated:now}:m));
    else { const code=genCode("Prep",prepList); setPrepList(p=>[...p,{id:Date.now(),code,name:form.name.trim(),unit:form.unit,class:form.class,yieldOverride:form.yieldOverride,ingredients:ings,lastUpdated:now}]); setShowCount(c=>c+1); }
    reset(); showToast(t.savedOk);
  };
  const doEdit=m=>{ setForm({name:m.name,unit:m.unit,class:m.class||"",yieldOverride:m.yieldOverride||"",ingredients:m.ingredients||[]}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setPrepList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };
  const addI=()=>setForm(f=>({...f,ingredients:[...f.ingredients,blank()]}));
  const remI=id=>setForm(f=>({...f,ingredients:f.ingredients.filter(i=>i.id!==id)}));
  const updI=(id,k,v)=>setForm(f=>({...f,ingredients:f.ingredients.map(i=>i.id===id?{...i,[k]:v}:i)}));
  const live=calcPrepCost({ingredients:form.ingredients,unit:form.unit,yieldOverride:form.yieldOverride});

  const doDownloadTemplate=()=>{
    const validClasses=cls.join(" / ")||"Sauce / Dough / Mix";
    const rows=[
      {"Prep Name":lang==="ar"?"دجاج متبل":"Marinated Chicken","Prep Code":"Prep-00001","Class":cls[0]||"Sauce","Unit":"kg","Ingredient Type":"Raw","Ingredient Code":"Raw-00001","Ingredient Name":lang==="ar"?"دجاج خام":"Raw Chicken","Qty (g/ml)":1000,"Waste %":20},
      {"Prep Name":lang==="ar"?"دجاج متبل":"Marinated Chicken","Prep Code":"Prep-00001","Class":cls[0]||"Sauce","Unit":"kg","Ingredient Type":"Raw","Ingredient Code":"Raw-00002","Ingredient Name":lang==="ar"?"بهارات":"Spices","Qty (g/ml)":30,"Waste %":0},
    ];
    const ws=XLSX.utils.json_to_sheet(rows); ws["!cols"]=[{wch:22},{wch:14},{wch:16},{wch:10},{wch:16},{wch:14},{wch:25},{wch:12},{wch:10}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Prep Items");
    XLSX.writeFile(wb,"TALA_Prep_Items_Template.xlsx");
  };
  const doExport=()=>{
    if(!prepList.length){doDownloadTemplate();return;}
    const rows=[];
    prepList.forEach(p=>{
      const ings=p.ingredients||[];
      if(!ings.length){rows.push({"Prep Name":p.name,"Prep Code":p.code,"Class":p.class||"","Unit":p.unit||"kg","Ingredient Type":"","Ingredient Code":"","Ingredient Name":"","Qty (g/ml)":"","Waste %":""});}
      else{ings.forEach(ing=>{
        let ingCode="",ingName="",ingType="";
        if(ing.source==="prep"){const pp=prepList.find(pp=>String(pp.id)===String(ing.rawId));if(pp){ingCode=pp.code;ingName=pp.name;ingType="Prep";}}
        else{const raw=rawList.find(r=>String(r.id)===String(ing.rawId));if(raw){ingCode=raw.code;ingName=raw.name;ingType="Raw";}}
        rows.push({"Prep Name":p.name,"Prep Code":p.code,"Class":p.class||"","Unit":p.unit||"kg","Ingredient Type":ingType,"Ingredient Code":ingCode,"Ingredient Name":ingName,"Qty (g/ml)":parseFloat(ing.qty)||0,"Waste %":parseFloat(ing.waste)||0});
      });}
    });
    const ws=XLSX.utils.json_to_sheet(rows); ws["!cols"]=[{wch:22},{wch:14},{wch:16},{wch:10},{wch:16},{wch:14},{wch:25},{wch:12},{wch:10}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Prep Items");
    XLSX.writeFile(wb,"TALA_Prep_Export.xlsx");
  };
  const doImport=e=>{
    const file=e.target.files[0]; if(!file) return;
    const r=new FileReader();
    r.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        if(!rows.length){showToast(t.noMatch,"warning");return;}
        const fv=(row,keys)=>{ for(const k of keys){const v=row[k];if(v!==undefined&&v!==null&&String(v).trim()!=="")return String(v).trim();}return ""; };
        const groups={};
        rows.forEach(row=>{
          const pn=fv(row,["Prep Name","اسم Prep","prep name","prep","اسم البريب","Prep","اسم"]);
          const pc=fv(row,["Prep Code","كود Prep","prep code"]);
          const pcls=fv(row,["Class","الكلاس","class","كلاس"]);
          const punit=fv(row,["Unit","الوحدة","unit"]);
          if(!pn||pn.startsWith("⚡")) return;
          const key=pn.toLowerCase();
          if(!groups[key]) groups[key]={code:pc,name:pn,class:pcls,unit:punit,ingredients:[]};
          else { if(pcls&&!groups[key].class)groups[key].class=pcls; }
          const ingCode=fv(row,["Ingredient Code","كود المكون","code"]);
          const ingName=fv(row,["Ingredient Name","اسم المكون","ingredient","المكون"]);
          const ingTypeRaw=fv(row,["Ingredient Type","نوع المكون","type"]);
          const ingType=ingTypeRaw.toLowerCase().includes("prep")||ingTypeRaw.includes("بريب")?"prep":"raw";
          const qtyRaw=fv(row,["Qty (g/ml)","الكمية (g/ml)","qty","quantity","الكمية","Qty"]);
          let qty=parseFloat(qtyRaw)||0;
          const wasteRaw=fv(row,["Waste %","الهدر %","waste"]);
          let waste=parseFloat(wasteRaw)||0;
          if(waste>0&&waste<1) waste=Math.round(waste*100*10)/10;
          if(!ingName&&!ingCode) return;
          if(ingType==="prep"){const p=prepList.find(p=>p.code===ingCode||p.name.toLowerCase()===ingName.toLowerCase());if(p) groups[key].ingredients.push({id:Date.now()+Math.random(),source:"prep",rawId:p.id,qty,waste});}
          else{const raw=rawList.find(r=>r.code===ingCode||r.name.toLowerCase()===ingName.toLowerCase());if(raw) groups[key].ingredients.push({id:Date.now()+Math.random(),source:"raw",rawId:raw.id,qty,waste});}
        });
        const items=Object.values(groups).map(pg=>{
          const unitRaw=(pg.unit||"").toLowerCase();
          const unit=["liter","litre","l","لتر"].includes(unitRaw)?"liter":["piece","pcs","حبة"].includes(unitRaw)?"piece":"kg";
          const existing=prepList.find(p=>p.name.toLowerCase()===pg.name.toLowerCase())||(pg.code?prepList.find(p=>p.code===pg.code):null);
          return {name:pg.name,code:pg.code,class:pg.class,unit,_ingredients:pg.ingredients,_existing:existing};
        });
        setPreviewData({items,module:"prep"});
        setShowImport(false);
      }catch(err){showToast(lang==="ar"?"خطأ في الملف":"File error","error");}
    };
    r.readAsBinaryString(file); e.target.value="";
  };
  const confirmImport=(editedItems)=>{
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    let updated=0,added=0;
    setPrepList(prev=>{
      const u=[...prev];
      editedItems.forEach(item=>{
        if(item._skip) return;
        const ings=item._ingredients||[];
        const idx=item._existing?u.findIndex(p=>p.id===item._existing.id):-1;
        if(idx!==-1){u[idx]={...u[idx],name:item.name.trim(),class:item.class||u[idx].class,unit:item.unit||u[idx].unit,ingredients:ings,lastUpdated:now};updated++;}
        else{
          const nameExists=u.some(m=>m.name.trim().toLowerCase()===item.name.trim().toLowerCase());
          if(nameExists) return;
          const newCode=item.code&&!["undefined","nan",""].includes(item.code)?item.code:genCode("Prep",u);
          u.push({id:Date.now()+Math.random(),code:newCode,name:item.name.trim(),class:item.class||"",unit:item.unit||"kg",yieldOverride:"",ingredients:ings,lastUpdated:now});
          added++;
        }
      });
      return u;
    });
    setPreviewData(null);
    showToast(lang==="ar"?`✅ أضيف ${added} | حُدِّث ${updated}`:`✅ Added ${added} | Updated ${updated}`,(added+updated)>0?"success":"warning");
  };
  const toggleSel=(id)=>setSelected(p=>{const n=new Set(p);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>setSelected(p=>p.size===filtered.length?new Set():new Set(filtered.map(m=>m.id)));
  const bulkDelete=()=>{setPrepList(p=>p.filter(m=>!selected.has(m.id)));setSelected(new Set());showToast(lang==="ar"?`تم حذف ${selected.size} بريب`:`Deleted ${selected.size} items`,"error");};

  // FIX: Search filter is strict — only exact class filter
  const filtered=prepList.filter(m=>{
    const matchSearch=!search||m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase());
    const matchClass=fcls==="all"||(m.class||"")===(fcls);
    return matchSearch&&matchClass;
  });
  const displayed=filtered.slice(0,showCount);
  const hasMore=filtered.length>showCount;

  return (
    <div>
      {dups.length>0&&<div style={{background:"#3d2205",border:"1px solid #ca8a04",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <span style={{fontSize:12,color:"#fcd34d",fontWeight:700}}>⚠️ {lang==="ar"?`يوجد ${dups.length} بريب بأسماء مكررة`:`${dups.length} duplicate prep names`}</span>
        <button onClick={()=>setBulkDelModal({items:dups,module:"prep"})} style={{background:"#ca8a04",color:"#060810",border:"none",borderRadius:6,padding:"5px 12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{lang==="ar"?"عرض وإدارة":"View & Manage"}</button>
      </div>}
      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:240}} placeholder={t.search} value={search} onChange={e=>{setSearch(e.target.value);setSelected(new Set());}}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={()=>setShowImport(true)}>{t.importXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>
      {selected.size>0&&hasPerm(mod,"delete")&&<div style={{background:"#1a0f0f",border:"1px solid #dc262666",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <span style={{fontSize:13,color:"#fca5a5",fontWeight:700}}>{lang==="ar"?`تم تحديد ${selected.size} بريب`:`${selected.size} selected`}</span>
        <div style={{display:"flex",gap:7}}>
          <button onClick={()=>setSelected(new Set())} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:6,padding:"5px 12px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{lang==="ar"?"إلغاء":"Deselect"}</button>
          <button onClick={bulkDelete} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"5px 14px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>{lang==="ar"?`حذف ${selected.size}`:`Delete ${selected.size}`}</button>
        </div>
      </div>}
      <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}}>
        <button className={`fbtn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all} ({prepList.length})</button>
        {cls.map(c=><button key={c} className={`fbtn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c} ({prepList.filter(m=>(m.class||"")===c).length})</button>)}
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{borderCollapse:"collapse"}}>
          <thead><tr>
            {hasPerm(mod,"delete")&&<th style={{width:36,padding:"10px 8px"}}><input type="checkbox" checked={filtered.length>0&&selected.size===filtered.length} onChange={toggleAll} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></th>}
            {["#",t.code,t.name,t.class,t.unit,lang==="ar"?"مكونات":"Ing.",t.yieldWeight,t.costPerUnit,t.usedInProducts,t.actions].map((h,i)=><th key={i}>{h}</th>)}
          </tr></thead>
          <tbody>
            {displayed.length===0?<tr><td colSpan={11} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :displayed.map((m,i)=>{
              const {costPerUnit,yieldKg}=calcPrepCost(m);
              const nProd=prodList.filter(p=>p.ingredients?.some(i=>i.source==="prep"&&String(i.srcId)===String(m.id))).length;
              const isDup=dups.some(d=>d.id===m.id);
              const isSel=selected.has(m.id);
              return <tr key={m.id} style={{background:isSel?C.accent+"12":isDup?"#3d220511":""}}>
                {hasPerm(mod,"delete")&&<td style={{padding:"10px 8px",textAlign:"center"}}><input type="checkbox" checked={isSel} onChange={()=>toggleSel(m.id)} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></td>}
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>{isDup&&<span style={{color:"#f59e0b",marginLeft:4,marginRight:4,fontSize:12}}>⚠</span>}{m.name}</td>
                <td><span className="badge badge-cls">{m.class||"—"}</span></td>
                <td><span className={`badge badge-${m.unit}`}>{unitLbl(m.unit,t)}</span></td>
                <td style={{color:C.muted}}>{m.ingredients?.length||0}</td>
                <td style={{color:C.muted}}>{yieldKg.toFixed(3)}</td>
                <td style={{color:C.accent,fontWeight:700}}>{costPerUnit.toFixed(4)}</td>
                <td>{nProd>0?<button onClick={()=>setUsageModal({item:m,type:"product",list:prodList.filter(p=>p.ingredients?.some(i=>i.source==="prep"&&String(i.srcId)===String(m.id)))})} style={{background:"#0a2a1a",color:"#4ade80",border:"1px solid #16a34a33",borderRadius:6,padding:"3px 10px",cursor:"pointer",fontWeight:700,fontSize:13}}>{nProd}</button>:<span style={{color:C.muted}}>0</span>}</td>
                <td><div style={{display:"flex",gap:5}}>
                  <button className="btn-sm-v" onClick={()=>setViewItem(m)}>{t.view}</button>
                  {hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>doEdit(m)}>{t.edit}</button>}
                  {hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>setDelId(m.id)}>{t.delete}</button>}
                </div></td>
              </tr>;
            })}
          </tbody>
        </table>
      </div></div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginTop:10}}>
        <span style={{fontSize:12,color:C.muted}}>{lang==="ar"?`عرض ${displayed.length} من ${filtered.length}`:`Showing ${displayed.length} of ${filtered.length}`}</span>
        {hasMore&&<button onClick={()=>setShowCount(c=>c+pageSize)} style={{background:C.surface,color:C.accent,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{t.showMore} (+{Math.min(pageSize,filtered.length-showCount)})</button>}
        {showCount>pageSize&&!hasMore&&<button onClick={()=>setShowCount(pageSize)} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{t.showLess}</button>}
      </div>

      {/* View modal */}
      {viewItem&&(()=>{
        const {costPerUnit,yieldKg,totalCost}=calcPrepCost(viewItem);
        return <div className="overlay" onClick={e=>e.target===e.currentTarget&&setViewItem(null)}>
          <div className="modal modal-lg" style={{maxWidth:820}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
              <div><div style={{fontWeight:800,fontSize:16,color:C.accent}}>{viewItem.name}</div><div style={{fontSize:12,color:C.muted,marginTop:2}}>{viewItem.code} · {viewItem.class||"—"} · {unitLbl(viewItem.unit,t)}</div></div>
              <button onClick={()=>setViewItem(null)} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer"}}>✕</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
              {[{l:lang==="ar"?"عدد المكونات":"Ingredients",v:viewItem.ingredients?.length||0,c:C.blue},{l:lang==="ar"?"وزن الناتج":"Yield",v:yieldKg.toFixed(3)+" "+unitLbl(viewItem.unit,t),c:C.green},{l:lang==="ar"?"تكلفة الباتش":"Batch Cost",v:totalCost.toFixed(2),c:"#f87171"},{l:lang==="ar"?"تكلفة/وحدة":"Cost/Unit",v:costPerUnit.toFixed(4),c:C.accent}].map((s,i)=>(
                <div key={i} style={{background:C.surface,borderRadius:10,padding:"12px 14px",border:`1px solid ${C.border}`}}><div style={{fontSize:10,color:C.muted,marginBottom:4,textTransform:"uppercase"}}>{s.l}</div><div style={{fontSize:17,fontWeight:800,color:s.c}}>{s.v}</div></div>
              ))}
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse"}}>
                <thead><tr>{["#",lang==="ar"?"المادة":"Material",lang==="ar"?"الكمية":"Qty",lang==="ar"?"الهدر %":"Waste %",lang==="ar"?"الكمية الصافية":"Net Qty",lang==="ar"?"سعر/وحدة":"Price/Unit",lang==="ar"?"التكلفة":"Cost"].map((h,i)=><th key={i} style={{padding:"9px 12px",textAlign:lang==="ar"?"right":"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {(viewItem.ingredients||[]).map((ing,i)=>{
                    const isPrep=ing.source==="prep";
                    const raw=!isPrep?rawList.find(r=>String(r.id)===String(ing.rawId)):null;
                    const prep=isPrep?prepList.find(p=>String(p.id)===String(ing.rawId||ing.srcId||ing.prepId)):null;
                    if(!raw&&!prep) return null;
                    const qty=parseFloat(ing.qty)||0; const waste=(parseFloat(ing.waste)||0)/100;
                    const netQty=qty*(1-waste);
                    let ingCost=0,ingName="",ingCode="",unit="g";
                    if(raw){ingName=raw.name;ingCode=raw.code;unit=raw.unit==="kg"?"g":raw.unit==="liter"?"ml":"pcs";ingCost=(raw.unit==="piece"?qty:qty/1000)*raw.price;}
                    else{const{costPerUnit}=calcPrepCost(prep);ingName=prep.name;ingCode=prep.code;unit=prep.unit==="kg"?"g":prep.unit==="liter"?"ml":"g";ingCost=(qty/1000)*costPerUnit;}
                    return <tr key={i} style={{borderBottom:`1px solid ${C.border}22`}}>
                      <td style={{padding:"10px 12px",color:C.muted,fontSize:11}}>{i+1}</td>
                      <td style={{padding:"10px 12px",fontWeight:600}}>{ingName} <span style={{color:C.muted,fontSize:11}}>({ingCode})</span>{isPrep&&<span style={{marginLeft:5,background:C.blue+"22",color:C.blue,padding:"1px 6px",borderRadius:10,fontSize:10,fontWeight:700}}>PREP</span>}</td>
                      <td style={{padding:"10px 12px"}}>{qty.toFixed(0)} {unit}</td>
                      <td style={{padding:"10px 12px"}}>{ing.waste>0?<span style={{background:C.yellow+"22",color:C.yellow,padding:"2px 8px",borderRadius:20,fontSize:12,fontWeight:700}}>{ing.waste}%</span>:<span style={{color:C.muted}}>0%</span>}</td>
                      <td style={{padding:"10px 12px",color:C.green,fontWeight:600}}>{netQty.toFixed(0)} {unit}</td>
                      <td style={{padding:"10px 12px",color:C.muted,fontSize:11}}>{raw?raw.price.toFixed(2)+"/"+unitLbl(raw.unit,t):"—"}</td>
                      <td style={{padding:"10px 12px",color:C.accent,fontWeight:700}}>{ingCost.toFixed(4)}</td>
                    </tr>;
                  })}
                  <tr style={{background:C.surface,borderTop:`2px solid ${C.border}`}}>
                    <td colSpan={6} style={{padding:"11px 12px",fontWeight:800}}>{lang==="ar"?"الإجمالي":"Total"}</td>
                    <td style={{padding:"11px 12px",fontWeight:800,color:"#f87171",fontSize:14}}>{totalCost.toFixed(4)}</td>
                  </tr>
                  <tr style={{background:C.accent+"12"}}>
                    <td colSpan={6} style={{padding:"10px 12px",fontWeight:700}}>{lang==="ar"?"تكلفة الوحدة":"Cost per unit"}</td>
                    <td style={{padding:"10px 12px",fontWeight:900,color:C.accent,fontSize:15}}>{costPerUnit.toFixed(4)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={{display:"flex",justifyContent:"flex-end",marginTop:14}}><button className="btn btn-secondary" onClick={()=>setViewItem(null)}>{t.cancel}</button></div>
          </div>
        </div>;
      })()}

      {showImport&&<ImportModal t={t} lang={lang} C={C} type={t.prepItem} onClose={()=>setShowImport(false)} onDownloadTemplate={doDownloadTemplate} onFileSelect={doImport}/>}
      {usageModal&&<UsageModal t={t} C={C} usageModal={usageModal} onClose={()=>setUsageModal(null)}/>}
      {previewData&&previewData.module==="prep"&&<PreviewModal t={t} lang={lang} C={C} data={previewData} onConfirm={confirmImport} onCancel={()=>setPreviewData(null)} classes={cls} moduleType="prep"/>}
      {bulkDelModal&&bulkDelModal.module==="prep"&&<DupDeleteModal t={t} lang={lang} C={C} items={bulkDelModal.items} onDelete={ids=>{setPrepList(p=>p.filter(m=>!ids.includes(m.id)));setBulkDelModal(null);showToast(lang==="ar"?`تم حذف ${ids.length} بريب مكرر`:`Deleted ${ids.length} duplicates`,"error");}} onClose={()=>setBulkDelModal(null)}/>}
      {showForm&&<div className="overlay" onClick={e=>{if(e.target===e.currentTarget)reset();}}><div className="modal modal-lg" onClick={e=>e.stopPropagation()}>
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
        {form.ingredients.map(ing=><IngRow key={ing.id} ing={ing} rawList={rawList} prepList={prepList} lang={lang} t={t} C={C} onUpdate={(k,v)=>updI(ing.id,k,v)} onRemove={()=>remI(ing.id)}/>)}
        <div style={{background:"#060810",borderRadius:8,padding:"10px 14px",marginTop:10,display:"flex",gap:20,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:C.muted}}>{t.yieldWeight}: <strong style={{color:C.green}}>{live.yieldKg.toFixed(3)} {form.unit}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>{lang==="ar"?"تكلفة الباتش":"Batch Cost"}: <strong style={{color:"#f87171"}}>{live.totalCost.toFixed(2)}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>{t.costPerUnit}: <strong style={{color:C.accent}}>{live.costPerUnit.toFixed(4)}</strong></span>
        </div>
        <div style={{display:"flex",gap:8,marginTop:14}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
      </div></div>}
      {delId&&<DelModal t={t} C={C} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════
// PRODUCTS TAB — with full Excel template (ingredients + prices + stdCost + date)
// ═══════════════════════════════════════════════════════════════
function ProductsTab({t,lang,C=DARK,prodList,setProdList,rawList,prepList,classes,calcPrepCost,calcProductCost,showToast,hasPerm,mod}) {
  const [searchRaw,setSearchRaw]=useState(""); const [search,setSearch]=useState("");
  const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",class:"",sellingPrice:"",posSellPrice:"",aggSellPrice:"",stdCost:"",ingredients:[]}); const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null); const [showImport,setShowImport]=useState(false);
  const [viewItem,setViewItem]=useState(null);
  const [pageSize,setPageSize]=useState(20); const [showCount,setShowCount]=useState(20);
  const [selected,setSelected]=useState(new Set());
  const [importDate,setImportDate]=useState(null);
  const fileRef=useRef();
  const cls=classes.products||[];
  const blank=()=>({id:Date.now()+Math.random(),source:"raw",srcId:"",qty:"",waste:"0",inputUnit:""});
  const reset=()=>{ setForm({name:"",class:"",sellingPrice:"",posSellPrice:"",aggSellPrice:"",stdCost:"",ingredients:[]}); setErrs({}); setShowForm(false); setEditId(null); };
  const ok=()=>{ const e={}; if(!form.name.trim()) e.name=t.required; setErrs(e); return !Object.keys(e).length; };
  const save=()=>{
    if(!ok()) return;
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    const ings=form.ingredients.filter(i=>i.srcId&&parseFloat(i.qty)>0);
    const posP=parseFloat(form.posSellPrice)||0;
    const aggP=parseFloat(form.aggSellPrice)||0;
    const sp=parseFloat(form.sellingPrice)||posP||aggP;
    const prod={name:form.name.trim(),class:form.class,sellingPrice:sp,posSellPrice:posP,aggSellPrice:aggP,stdCost:parseFloat(form.stdCost)||0,ingredients:ings,lastUpdated:now};
    if(editId!==null) setProdList(p=>p.map(m=>m.id===editId?{...m,...prod}:m));
    else { const code=genCode("Prod",prodList); setProdList(p=>[...p,{id:Date.now(),code,...prod}]); setShowCount(c=>c+1); }
    reset(); showToast(t.savedOk);
  };
  const doEdit=m=>{ setForm({name:m.name,class:m.class||"",sellingPrice:String(m.sellingPrice||""),posSellPrice:String(m.posSellPrice||""),aggSellPrice:String(m.aggSellPrice||""),stdCost:String(m.stdCost||""),ingredients:m.ingredients||[]}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setProdList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };
  const addI=()=>setForm(f=>({...f,ingredients:[...f.ingredients,blank()]}));
  const remI=id=>setForm(f=>({...f,ingredients:f.ingredients.filter(i=>i.id!==id)}));
  const updI=(id,k,v)=>setForm(f=>({...f,ingredients:f.ingredients.map(i=>i.id===id?{...i,[k]:v}:i)}));
  const live=useMemo(()=>calcProductCost({ingredients:form.ingredients,posSellPrice:form.posSellPrice,aggSellPrice:form.aggSellPrice,sellingPrice:form.sellingPrice}),[form.ingredients,form.posSellPrice,form.aggSellPrice,form.sellingPrice,calcProductCost]);

  const doDownloadTemplate=()=>{
    // Full template with ingredients + prices
    const rows=[
      {"Product Name":lang==="ar"?"سلطة تالا":"Tala Salad","Product Code":"Prod-00001","Class":cls[0]||"Main Dish","POS Price":35,"AGG Price":40,"Std Cost":10,"Ingredient Source":"Raw","Ingredient Code":"Raw-00001","Ingredient Name":lang==="ar"?"خس":"Lettuce","Qty (g)":150,"Waste %":10},
      {"Product Name":lang==="ar"?"سلطة تالا":"Tala Salad","Product Code":"Prod-00001","Class":"","POS Price":"","AGG Price":"","Std Cost":"","Ingredient Source":"Raw","Ingredient Code":"Raw-00002","Ingredient Name":lang==="ar"?"زيت زيتون":"Olive Oil","Qty (g)":20,"Waste %":0},
    ];
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=[{wch:22},{wch:14},{wch:14},{wch:12},{wch:12},{wch:12},{wch:16},{wch:14},{wch:25},{wch:10},{wch:10}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Products");
    XLSX.writeFile(wb,"TALA_Products_Template.xlsx");
  };
  const doExport=()=>{
    if(!prodList.length){doDownloadTemplate();return;}
    const rows=[];
    prodList.forEach(p=>{
      const ings=p.ingredients||[];
      if(!ings.length){rows.push({"Product Name":p.name,"Product Code":p.code,"Class":p.class||"","POS Price":p.posSellPrice||"","AGG Price":p.aggSellPrice||"","Std Cost":p.stdCost||"","Ingredient Source":"","Ingredient Code":"","Ingredient Name":"","Qty (g)":"","Waste %":""});}
      else{ings.forEach((ing,idx)=>{
        let ingCode="",ingName="",ingSrc="";
        if(ing.source==="prep"){const pp=prepList.find(pp=>String(pp.id)===String(ing.srcId));if(pp){ingCode=pp.code;ingName=pp.name;ingSrc="Prep";}}
        else{const raw=rawList.find(r=>String(r.id)===String(ing.srcId));if(raw){ingCode=raw.code;ingName=raw.name;ingSrc="Raw";}}
        rows.push({"Product Name":idx===0?p.name:"","Product Code":idx===0?p.code:"","Class":idx===0?(p.class||""):"","POS Price":idx===0?(p.posSellPrice||""):"","AGG Price":idx===0?(p.aggSellPrice||""):"","Std Cost":idx===0?(p.stdCost||""):"","Ingredient Source":ingSrc,"Ingredient Code":ingCode,"Ingredient Name":ingName,"Qty (g)":parseFloat(ing.qty)||0,"Waste %":parseFloat(ing.waste)||0});
      });}
    });
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=[{wch:22},{wch:14},{wch:14},{wch:12},{wch:12},{wch:12},{wch:16},{wch:14},{wch:25},{wch:10},{wch:10}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Products");
    XLSX.writeFile(wb,"TALA_Products_Export.xlsx");
  };
  const doImport=e=>{
    const file=e.target.files[0]; if(!file) return;
    const r=new FileReader();
    r.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        if(!rows.length){showToast(t.noMatch,"warning");return;}
        const fv=(row,keys)=>{ for(const k of keys){const v=row[k];if(v!==undefined&&v!==null&&String(v).trim()!=="")return String(v).trim();}return ""; };
        const groups={};
        rows.forEach(row=>{
          const pn=fv(row,["Product Name","اسم المنتج","product name","Product","product"]);
          const pc=fv(row,["Product Code","كود المنتج","product code"]);
          const pcls=fv(row,["Class","الكلاس","class"]);
          const posP=parseFloat(fv(row,["POS Price","سعر المحل","pos price","pos"]))||0;
          const aggP=parseFloat(fv(row,["AGG Price","سعر التطبيقات","agg price","agg"]))||0;
          const stdC=parseFloat(fv(row,["Std Cost","التكلفة المعيارية","std cost","standard cost"]))||0;
          if(!pn) return;
          const key=pn.toLowerCase();
          if(!groups[key]) groups[key]={code:pc,name:pn,class:pcls,posP,aggP,stdC,ingredients:[]};
          const ingCode=fv(row,["Ingredient Code","كود المكون","ingredient code"]);
          const ingName=fv(row,["Ingredient Name","اسم المكون","ingredient name","ingredient"]);
          const ingSrcRaw=fv(row,["Ingredient Source","مصدر المكون","source","ingredient type"]);
          const ingSrc=ingSrcRaw.toLowerCase().includes("prep")?"prep":"raw";
          const qty=parseFloat(fv(row,["Qty (g)","الكمية (g)","qty","quantity","الكمية"]))||0;
          const waste=parseFloat(fv(row,["Waste %","الهدر %","waste"]))||0;
          if(!ingName&&!ingCode) return;
          let srcId=null;
          if(ingSrc==="prep"){const p=prepList.find(p=>p.code===ingCode||p.name.toLowerCase()===ingName.toLowerCase());if(p)srcId=p.id;}
          else{const raw=rawList.find(r=>r.code===ingCode||r.name.toLowerCase()===ingName.toLowerCase());if(raw)srcId=raw.id;}
          if(srcId) groups[key].ingredients.push({id:Date.now()+Math.random(),source:ingSrc,srcId,qty,waste:""+waste,inputUnit:""});
        });
        const items=Object.values(groups).map(pg=>{
          const existing=prodList.find(p=>p.name.toLowerCase()===pg.name.toLowerCase())||(pg.code?prodList.find(p=>p.code===pg.code):null);
          return {name:pg.name,code:pg.code,class:pg.class,posSellPrice:pg.posP,aggSellPrice:pg.aggP,stdCost:pg.stdC,_ingredients:pg.ingredients,_existing:existing};
        });
        // Direct import without preview for products (since ingredients don't fit preview modal)
        const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
        const dateLabel=importDate?`${mLabel(importDate.month,lang)} ${importDate.year}`:now;
        let added=0,updated=0;
        setProdList(prev=>{
          const u=[...prev];
          items.forEach(item=>{
            const idx=item._existing?u.findIndex(p=>p.id===item._existing.id):-1;
            const rec={name:item.name.trim(),class:item.class||"",posSellPrice:parseFloat(item.posSellPrice)||0,aggSellPrice:parseFloat(item.aggSellPrice)||0,stdCost:parseFloat(item.stdCost)||0,sellingPrice:parseFloat(item.posSellPrice)||0,ingredients:item._ingredients||[],lastUpdated:dateLabel};
            if(idx!==-1){u[idx]={...u[idx],...rec};updated++;}
            else{
              const nameExists=u.some(m=>m.name.trim().toLowerCase()===item.name.trim().toLowerCase());
              if(nameExists) return;
              const newCode=item.code&&!["undefined","nan",""].includes(item.code)?item.code:genCode("Prod",u);
              u.push({id:Date.now()+Math.random(),code:newCode,...rec});added++;
            }
          });
          return u;
        });
        setImportDate(null);
        showToast(lang==="ar"?`✅ أضيف ${added} | حُدِّث ${updated}`:`✅ Added ${added} | Updated ${updated}`,(added+updated)>0?"success":"warning");
      }catch(err){showToast(lang==="ar"?"خطأ في الملف":"File error","error");}
    };
    r.readAsBinaryString(file); e.target.value="";
  };

  const toggleSel=(id)=>setSelected(p=>{const n=new Set(p);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>setSelected(p=>p.size===displayed.length?new Set():new Set(displayed.map(m=>m.id)));
  const bulkDelete=()=>{setProdList(p=>p.filter(m=>!selected.has(m.id)));setSelected(new Set());showToast(lang==="ar"?`تم حذف ${selected.size} منتج`:`Deleted ${selected.size} products`,"error");};

  useEffect(()=>{const timer=setTimeout(()=>setSearch(searchRaw),200);return()=>clearTimeout(timer);},[searchRaw]);
  const filtered=prodList.filter(m=>(m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase()))&&(fcls==="all"||(m.class||"")===fcls));
  const displayed=filtered.slice(0,showCount);
  const hasMore=filtered.length>showCount;

  return (
    <div>
      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:220}} placeholder={t.search} value={searchRaw} onChange={e=>{setSearchRaw(e.target.value);setSelected(new Set());}}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{display:"flex",alignItems:"center",gap:5,background:C.surface,borderRadius:7,padding:"4px 8px",border:`1px solid ${C.border}`}}>
            <span style={{fontSize:11,color:C.muted}}>{t.pageSize}:</span>
            {[20,50,100,200,300].map(n=><button key={n} onClick={()=>{setPageSize(n);setShowCount(n);}} style={{background:pageSize===n?C.accent:"transparent",color:pageSize===n?"#080b14":C.muted,border:"none",borderRadius:5,padding:"3px 8px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{n}</button>)}
          </div>
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={()=>setShowImport(true)}>{t.importXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>
      {selected.size>0&&hasPerm(mod,"delete")&&<div style={{background:"#1a0f0f",border:"1px solid #dc262666",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <span style={{fontSize:13,color:"#fca5a5",fontWeight:700}}>{lang==="ar"?`تم تحديد ${selected.size} منتج`:`${selected.size} selected`}</span>
        <div style={{display:"flex",gap:7}}>
          <button onClick={()=>setSelected(new Set())} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:6,padding:"5px 12px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{lang==="ar"?"إلغاء":"Deselect"}</button>
          <button onClick={bulkDelete} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"5px 14px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>{lang==="ar"?`حذف ${selected.size}`:`Delete ${selected.size}`}</button>
        </div>
      </div>}
      <div style={{display:"flex",gap:7,marginBottom:10,flexWrap:"wrap"}}>
        <button className={`filter-btn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all} ({prodList.length})</button>
        {cls.map(c=><button key={c} className={`filter-btn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c} ({prodList.filter(m=>(m.class||"")===c).length})</button>)}
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{borderCollapse:"collapse"}}>
          <thead><tr>
            {hasPerm(mod,"delete")&&<th style={{width:36,padding:"10px 8px"}}><input type="checkbox" checked={displayed.length>0&&selected.size===displayed.length} onChange={toggleAll} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></th>}
            {["#",t.code,t.productName,t.class,lang==="ar"?"مكونات":"Ing.",t.totalCost,t.stdCost,"POS","POS%","AGG","AGG%",t.actions].map((h,i)=><th key={i}>{h}</th>)}
          </tr></thead>
          <tbody>
            {displayed.length===0?<tr><td colSpan={13} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :displayed.map((m,i)=>{
              const {totalCost,posMargin,aggMargin}=calcProductCost(m);
              const isSel=selected.has(m.id);
              const posP=parseFloat(m.posSellPrice||m.sellingPrice||0);
              const aggP=parseFloat(m.aggSellPrice||m.sellingPrice||0);
              return <tr key={m.id} style={{background:isSel?C.accent+"12":""}}>
                {hasPerm(mod,"delete")&&<td style={{padding:"10px 8px",textAlign:"center"}}><input type="checkbox" checked={isSel} onChange={()=>toggleSel(m.id)} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></td>}
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>{m.name}</td>
                <td><span className="badge badge-cls">{m.class||"—"}</span></td>
                <td style={{color:C.muted}}>{m.ingredients?.length||0}</td>
                <td style={{color:"#f87171",fontWeight:600}}>{totalCost.toFixed(2)}</td>
                <td style={{color:C.muted,fontWeight:500}}>{m.stdCost>0?parseFloat(m.stdCost).toFixed(2):"—"}</td>
                <td style={{color:"#4ade80",fontWeight:600}}>{posP>0?posP.toFixed(2):"—"}</td>
                <td><span style={{color:posP>0?(posMargin>30?"#4ade80":posMargin>15?"#fbbf24":"#f87171"):"#5a6585",fontWeight:700}}>{posP>0?posMargin.toFixed(1)+"%":"—"}</span></td>
                <td style={{color:"#60a5fa",fontWeight:600}}>{aggP>0?aggP.toFixed(2):"—"}</td>
                <td><span style={{color:aggP>0?(aggMargin>30?"#4ade80":aggMargin>15?"#fbbf24":"#f87171"):"#5a6585",fontWeight:700}}>{aggP>0?aggMargin.toFixed(1)+"%":"—"}</span></td>
                <td><div style={{display:"flex",gap:5}}>
                  <button className="btn-sm-v" onClick={()=>setViewItem(m)}>{t.view}</button>
                  {hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>doEdit(m)}>{t.edit}</button>}
                  {hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>setDelId(m.id)}>{t.delete}</button>}
                </div></td>
              </tr>;
            })}
          </tbody>
        </table>
      </div></div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginTop:10}}>
        <span style={{fontSize:12,color:C.muted}}>{lang==="ar"?`عرض ${displayed.length} من ${filtered.length}`:`Showing ${displayed.length} of ${filtered.length}`}</span>
        {hasMore&&<button onClick={()=>setShowCount(c=>c+pageSize)} style={{background:C.surface,color:C.accent,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{t.showMore} (+{Math.min(pageSize,filtered.length-showCount)})</button>}
        {showCount>pageSize&&!hasMore&&<button onClick={()=>setShowCount(pageSize)} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{t.showLess}</button>}
      </div>

      {/* View Product Modal */}
      {viewItem&&(()=>{
        const {totalCost,posMargin,aggMargin}=calcProductCost(viewItem);
        const posP=parseFloat(viewItem.posSellPrice||viewItem.sellingPrice||0);
        const aggP=parseFloat(viewItem.aggSellPrice||viewItem.sellingPrice||0);
        return <div className="overlay" onClick={e=>e.target===e.currentTarget&&setViewItem(null)}>
          <div className="modal modal-lg" style={{maxWidth:860}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
              <div><div style={{fontWeight:800,fontSize:16,color:C.accent}}>{viewItem.name}</div><div style={{fontSize:12,color:C.muted,marginTop:2}}>{viewItem.code} · {viewItem.class||"—"}</div></div>
              <button onClick={()=>setViewItem(null)} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer"}}>✕</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:18}}>
              {[
                {l:lang==="ar"?"التكلفة الفعلية":"Actual Cost",v:totalCost.toFixed(2),c:"#f87171"},
                {l:lang==="ar"?"التكلفة المعيارية":"Std Cost",v:parseFloat(viewItem.stdCost)>0?parseFloat(viewItem.stdCost).toFixed(2):"—",c:C.muted},
                {l:"POS "+t.sellingPrice,v:posP>0?posP.toFixed(2):"—",c:"#4ade80"},
                {l:"POS "+t.margin,v:posP>0?posMargin.toFixed(1)+"%":"—",c:posMargin>30?"#4ade80":posMargin>15?"#fbbf24":"#f87171"},
                {l:"AGG "+t.margin,v:aggP>0?aggMargin.toFixed(1)+"%":"—",c:aggMargin>30?"#4ade80":aggMargin>15?"#fbbf24":"#f87171"},
              ].map((s,i)=><div key={i} style={{background:C.surface,borderRadius:10,padding:"12px 14px",border:`1px solid ${C.border}`}}><div style={{fontSize:10,color:C.muted,marginBottom:4,textTransform:"uppercase"}}>{s.l}</div><div style={{fontSize:17,fontWeight:800,color:s.c}}>{s.v}</div></div>)}
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse"}}>
                <thead><tr>{["#",lang==="ar"?"المادة":"Material",lang==="ar"?"النوع":"Type",lang==="ar"?"الكمية (g)":"Qty (g)",lang==="ar"?"الهدر %":"Waste %",lang==="ar"?"التكلفة":"Cost"].map((h,i)=><th key={i} style={{padding:"9px 12px",textAlign:lang==="ar"?"right":"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {(viewItem.ingredients||[]).map((ing,i)=>{
                    const qty=parseFloat(ing.qty)||0; const waste=(parseFloat(ing.waste)||0)/100;
                    let name="—",cost=0,type="";
                    if(ing.source==="raw"){const r=rawList.find(r=>String(r.id)===String(ing.srcId));if(r){name=r.name+" ("+r.code+")";cost=(r.unit==="piece"?qty:qty/1000)*r.price;type="Raw";}}
                    else{const p=prepList.find(p=>String(p.id)===String(ing.srcId));if(p){name=p.name+" ("+p.code+")";const {costPerUnit}=calcPrepCost(p);cost=(p.unit==="piece"?qty:qty/1000)*costPerUnit;type="Prep";}}
                    return <tr key={i} style={{borderBottom:`1px solid ${C.border}22`}}>
                      <td style={{padding:"10px 12px",color:C.muted,fontSize:11}}>{i+1}</td>
                      <td style={{padding:"10px 12px",fontWeight:600}}>{name}</td>
                      <td style={{padding:"10px 12px"}}><span style={{background:type==="Raw"?"#0f204077":"#1a0f3377",color:type==="Raw"?"#60a5fa":"#a78bfa",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700}}>{type}</span></td>
                      <td style={{padding:"10px 12px"}}>{qty.toFixed(0)} g</td>
                      <td style={{padding:"10px 12px"}}>{parseFloat(ing.waste)>0?<span style={{background:C.yellow+"22",color:C.yellow,padding:"2px 8px",borderRadius:20,fontSize:12,fontWeight:700}}>{ing.waste}%</span>:<span style={{color:C.muted}}>0%</span>}</td>
                      <td style={{padding:"10px 12px",color:C.accent,fontWeight:700}}>{cost.toFixed(4)}</td>
                    </tr>;
                  })}
                  <tr style={{background:C.accent+"12"}}>
                    <td colSpan={5} style={{padding:"11px 12px",fontWeight:800}}>{lang==="ar"?"إجمالي التكلفة":"Total Cost"}</td>
                    <td style={{padding:"11px 12px",fontWeight:900,color:"#f87171",fontSize:15}}>{totalCost.toFixed(4)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={{display:"flex",justifyContent:"flex-end",marginTop:14,gap:8}}>
              {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{setViewItem(null);doEdit(viewItem);}}>{t.edit}</button>}
              <button className="btn btn-secondary" onClick={()=>setViewItem(null)}>{t.cancel}</button>
            </div>
          </div>
        </div>;
      })()}

      {showImport&&<ImportModal t={t} lang={lang} C={C} type={t.products} onClose={()=>setShowImport(false)} onDownloadTemplate={doDownloadTemplate} onFileSelect={doImport} showDatePicker={true} onDateSelect={setImportDate}/>}
      {showForm&&<div className="overlay" onClick={e=>{if(e.target===e.currentTarget)reset();}}><div className="modal modal-lg" onClick={e=>e.stopPropagation()}>
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.edit:t.add} — {t.products}</h2>
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:10,marginBottom:10}}>
          <div><label className="lbl">{t.productName}</label><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">{t.class}</label><select value={form.class} onChange={e=>setForm({...form,class:e.target.value})}><option value="">—</option>{cls.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10,background:C.surface,borderRadius:9,padding:"12px 14px",border:`1px solid ${C.border}`}}>
          <div>
            <label className="lbl" style={{color:"#22c55e"}}>{t.posSellPrice}</label>
            <input type="number" min="0" step="0.01" value={form.posSellPrice} onChange={e=>setForm({...form,posSellPrice:e.target.value})} placeholder="0.00"/>
            {parseFloat(form.posSellPrice)>0&&<div style={{fontSize:11,color:"#22c55e",marginTop:3}}>{lang==="ar"?"هامش:":"Margin:"} {live.posMargin.toFixed(1)}%</div>}
          </div>
          <div>
            <label className="lbl" style={{color:"#3b82f6"}}>{t.aggSellPrice}</label>
            <input type="number" min="0" step="0.01" value={form.aggSellPrice} onChange={e=>setForm({...form,aggSellPrice:e.target.value})} placeholder="0.00"/>
            {parseFloat(form.aggSellPrice)>0&&<div style={{fontSize:11,color:"#3b82f6",marginTop:3}}>{lang==="ar"?"هامش:":"Margin:"} {live.aggMargin.toFixed(1)}%</div>}
          </div>
          <div>
            <label className="lbl">{t.stdCost}</label>
            <input type="number" min="0" step="0.01" value={form.stdCost||""} onChange={e=>setForm({...form,stdCost:e.target.value})} placeholder="0.00"/>
          </div>
        </div>
        <div className="divider"/>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <span style={{fontWeight:700,fontSize:13}}>{t.ingredients}</span>
          <button className="btn btn-secondary" style={{padding:"5px 12px",fontSize:12}} onClick={addI}>+ {t.addIngredient}</button>
        </div>
        {form.ingredients.map(ing=><IngRowProd key={ing.id} ing={ing} rawList={rawList} prepList={prepList} lang={lang} t={t} C={C} onUpdate={(field,val)=>updI(ing.id,field,val)} onRemove={()=>remI(ing.id)}/>)}
        <div style={{background:C.bg||"#060810",borderRadius:8,padding:"10px 13px",marginTop:8,display:"flex",gap:16,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:C.muted}}>{t.totalCost}: <strong style={{color:C.red}}>{live.totalCost.toFixed(4)}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>POS: <strong style={{color:"#22c55e"}}>{parseFloat(form.posSellPrice||0).toFixed(2)}</strong> → <strong style={{color:live.posMargin>30?"#22c55e":live.posMargin>15?"#fbbf24":"#f87171"}}>{live.posMargin.toFixed(1)}%</strong></span>
          <span style={{fontSize:12,color:C.muted}}>AGG: <strong style={{color:"#60a5fa"}}>{parseFloat(form.aggSellPrice||0).toFixed(2)}</strong> → <strong style={{color:live.aggMargin>30?"#22c55e":live.aggMargin>15?"#fbbf24":"#f87171"}}>{live.aggMargin.toFixed(1)}%</strong></span>
        </div>
        <div style={{display:"flex",gap:8,marginTop:16}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
      </div></div>}
      {delId&&<DelModal t={t} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}


// ==================== MODIFIERS TAB ====================
function ModifiersTab({t,lang,C=DARK,mod,rawList=[],prepList=[],classes:clsList,hasPerm,modList,setModList,calcPrepCost,showToast,setClasses}){
  const SK_M="tc_mods_v1";
  const [importDate,setImportDate]=useState({month:String(new Date().getMonth()+1).padStart(2,"0"),year:String(new Date().getFullYear())});
  const [showForm,setShowForm]=useState(false);
  const [editId,setEditId]=useState(null);
  const [delId,setDelId]=useState(null);
  const [viewItem,setViewItem]=useState(null);
  const [showImport,setShowImport]=useState(false);
  const [filterName,setFilterName]=useState("");
  const [filterClass,setFilterClass]=useState("");
  const [page,setPage]=useState(1);
  const PAGE_SIZE=20;
  const blank={name:"",code:"",class:"",addOnPrice:"",ingredients:[]};
  const [form,setForm]=useState(blank);
  const [errs,setErrs]=useState({});

  const list=modList||[];
  const persist=arr=>{setModList(arr);localStorage.setItem(SK_M,JSON.stringify(arr));};

  const nextCode=()=>{
    const nums=list.map(r=>parseInt(r.code||0)).filter(n=>!isNaN(n));
    return String(nums.length?Math.max(...nums)+1:1).padStart(3,"0");
  };

  const calcCost=ings=>{
    if(!ings||!ings.length)return 0;
    return ings.reduce((s,ing)=>{
      const qty=parseFloat(ing.qty)||0;
      const waste=(parseFloat(ing.waste)||0)/100;
      const netQty=qty*(1-waste);
      if(ing.source==="prep"||ing.sourceType==="prep"){
        const p=prepList.find(x=>String(x.id)===String(ing.srcId)||x.code===ing.code);
        if(!p)return s;
        if(calcPrepCost){const {costPerUnit}=calcPrepCost(p);return s+(p.unit==="piece"?netQty:netQty/1000)*costPerUnit;}
        return s;
      } else {
        const r=rawList.find(x=>String(x.id)===String(ing.srcId)||x.code===ing.code);
        if(!r)return s;
        return s+(r.unit==="piece"?netQty:netQty/1000)*parseFloat(r.price||0);
      }
    },0);
  };

  const live=useMemo(()=>{
    const cost=calcCost(form.ingredients);
    const price=parseFloat(form.addOnPrice)||0;
    const margin=price>0?((price-cost)/price)*100:0;
    return {cost,margin};
  },[form,rawList,prepList]);

  const addI=()=>setForm(f=>({...f,ingredients:[...f.ingredients,{id:Date.now()+Math.random(),source:"raw",srcId:"",qty:"1",inputUnit:"g",waste:"0"}]}));
  const updI=(id,field,val)=>setForm(f=>({...f,ingredients:f.ingredients.map(i=>i.id===id?{...i,[field]:val}:i)}));
  const remI=id=>setForm(f=>({...f,ingredients:f.ingredients.filter(i=>i.id!==id)}));

  const validate=()=>{
    const e={};
    if(!form.name.trim())e.name=lang==="ar"?"مطلوب":"Required";
    const dupName=list.find(r=>r.name.trim().toLowerCase()===form.name.trim().toLowerCase()&&r.id!==editId);
    if(dupName)e.name=lang==="ar"?"الاسم مكرر":"Duplicate name";
    if(!form.addOnPrice||isNaN(parseFloat(form.addOnPrice)))e.addOnPrice=lang==="ar"?"مطلوب":"Required";
    setErrs(e);
    return !Object.keys(e).length;
  };

  const save=()=>{
    if(!validate())return;
    const code=form.code||nextCode();
    if(editId){
      persist(list.map(r=>r.id===editId?{...form,code,id:editId}:r));
    } else {
      persist([...list,{...form,code,id:Date.now()}]);
    }
    showToast&&showToast(lang==="ar"?"تم الحفظ":"Saved");
    reset();
  };

  const reset=()=>{setForm(blank);setEditId(null);setShowForm(false);setErrs({});};
  const doEdit=item=>{setForm({...item});setEditId(item.id);setShowForm(true);};
  const doDelete=id=>{persist(list.filter(r=>r.id!==id));setDelId(null);showToast&&showToast(lang==="ar"?"تم الحذف":"Deleted","red");};

  const doDownloadTemplate=()=>{
    const headers=["code","name","class","addOnPrice","ingredientSource","ingredientCode","ingredientName","qty_g","waste_pct"];
    const example=["M001","Cheese Extra","Modifiers","2.50","raw","RM001","Mozzarella","30","5"];
    const ws=XLSX.utils.aoa_to_sheet([headers,example]);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Modifiers");
    XLSX.writeFile(wb,"modifiers_template.xlsx");
  };

  const doImport=file=>{
    const reader=new FileReader();
    reader.onload=e=>{
      const wb=XLSX.read(e.target.result,{type:"binary"});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      let added=0,dupes=0;
      const updated=[...list];
      // Group rows by modifier name/code
      const groups={};
      rows.forEach(row=>{
        const name=(row.name||row.Name||"").toString().trim();
        const code=(row.code||row.Code||"").toString().trim();
        const key=code||name;
        if(!key)return;
        if(!groups[key])groups[key]={name,code,class:row.class||row.Class||"",addOnPrice:row.addOnPrice||row.add_on_price||0,ingredients:[]};
        const ingSrc=(row.ingredientSource||row.ingredient_source||"").toString().trim().toLowerCase();
        const ingCode=(row.ingredientCode||row.ingredient_code||"").toString().trim();
        const ingName=(row.ingredientName||row.ingredient_name||"").toString().trim();
        const qty=parseFloat(row.qty_g||row.qty||1);
        const waste=parseFloat(row.waste_pct||0);
        if(ingCode||ingName){
          const srcType=ingSrc==="prep"?"prep":"raw";
          const mat=srcType==="raw"?rawList.find(r=>r.code===ingCode||r.name===ingName):prepList.find(p=>p.code===ingCode||p.name===ingName);
          groups[key].ingredients.push({id:Date.now()+Math.random(),sourceType:srcType,code:mat?.code||ingCode,name:mat?.name||ingName,qty:String(qty),inputUnit:"g",unit:"g",wastePct:waste});
        }
      });
      Object.values(groups).forEach(g=>{
        const n=g.name;
        const dup=updated.find(r=>r.name.toLowerCase()===n.toLowerCase());
        if(dup){dupes++;return;}
        updated.push({id:Date.now()+Math.random(),code:g.code||nextCode(),...g});
        added++;
      });
      persist(updated);
      showToast&&showToast(`${lang==="ar"?"مُضاف:":"Added:"} ${added} | ${lang==="ar"?"مكرر:":"Dupes:"} ${dupes}`);
      setShowImport(false);
    };
    reader.readAsBinaryString(file);
  };

  const doExport=()=>{
    const rows=[];
    list.forEach(r=>{
      const cost=calcCost(r.ingredients);
      const price=parseFloat(r.addOnPrice)||0;
      const margin=price>0?((price-cost)/price)*100:0;
      if(!r.ingredients||!r.ingredients.length){
        rows.push({code:r.code,name:r.name,class:r.class||"",addOnPrice:r.addOnPrice,cost:cost.toFixed(4),margin:margin.toFixed(1)+"%",ingredientSource:"",ingredientCode:"",ingredientName:"",qty_g:"",waste_pct:""});
      } else {
        r.ingredients.forEach((ing,idx)=>{
          rows.push({
            code:idx===0?r.code:"",name:idx===0?r.name:"",class:idx===0?r.class||"":"",addOnPrice:idx===0?r.addOnPrice:"",cost:idx===0?cost.toFixed(4):"",margin:idx===0?margin.toFixed(1)+"%":"",
            ingredientSource:ing.sourceType,ingredientCode:ing.code,ingredientName:ing.name,qty_g:ing.qty,waste_pct:ing.wastePct||0
          });
        });
      }
    });
    const ws=XLSX.utils.json_to_sheet(rows);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Modifiers");
    XLSX.writeFile(wb,"modifiers_export.xlsx");
  };

  const allClasses=(clsList?.modifiers||[]).concat(list.map(r=>r.class).filter(Boolean).filter(c=>!(clsList?.modifiers||[]).includes(c))).sort();

  const filtered=useMemo(()=>list.filter(r=>{
    if(filterClass&&r.class!==filterClass)return false;
    if(filterName){
      const q=filterName.toLowerCase();
      if(!r.name.toLowerCase().startsWith(q)&&!r.code?.toLowerCase().startsWith(q)&&!r.name.toLowerCase().includes(q))return false;
    }
    return true;
  }),[list,filterName,filterClass]);

  const totalPages=Math.max(1,Math.ceil(filtered.length/PAGE_SIZE));
  const paged=filtered.slice((page-1)*PAGE_SIZE,page*PAGE_SIZE);

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:8}}>
        <h2 style={{fontSize:15,fontWeight:700,color:C.accent,margin:0}}>{t.modifiers||"Modifiers"} <span style={{fontSize:12,color:C.muted}}>({list.length})</span></h2>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {hasPerm(mod,"import")&&<button className="btn btn-secondary" style={{fontSize:12}} onClick={()=>setShowImport(true)}>{t.import}</button>}
          {hasPerm(mod,"export")&&<button className="btn btn-secondary" style={{fontSize:12}} onClick={doExport}>{t.export}</button>}
          {hasPerm(mod,"add")&&<button className="btn btn-primary" style={{fontSize:12}} onClick={()=>{setForm({...blank,code:nextCode()});setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>

      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
        <input placeholder={lang==="ar"?"بحث باسم/كود":"Search name/code"} value={filterName} onChange={e=>{setFilterName(e.target.value);setPage(1);}} style={{flex:1,minWidth:120}}/>
        <select value={filterClass} onChange={e=>{setFilterClass(e.target.value);setPage(1);}} style={{minWidth:100}}>
          <option value="">{lang==="ar"?"كل الفئات":"All Classes"}</option>
          {allClasses.map(c=><option key={c} value={c}>{c}</option>)}
        </select>
        {(filterName||filterClass)&&<button className="btn btn-secondary" style={{fontSize:11}} onClick={()=>{setFilterName("");setFilterClass("");setPage(1);}}>✕</button>}
      </div>

      <div className="table-wrap">
        <table>
          <thead><tr>
            <th>{t.code}</th><th>{t.name}</th><th>{t.class}</th>
            <th>{lang==="ar"?"سعر الإضافة":"Add-on Price"}</th>
            <th>{t.cost||"Cost"}</th>
            <th>{lang==="ar"?"الهامش":"Margin"}%</th>
            <th>{t.actions}</th>
          </tr></thead>
          <tbody>
            {paged.length===0&&<tr><td colSpan={7} style={{textAlign:"center",color:C.muted,padding:20}}>{t.noData}</td></tr>}
            {paged.map(r=>{
              const cost=calcCost(r.ingredients);
              const price=parseFloat(r.addOnPrice)||0;
              const margin=price>0?((price-cost)/price)*100:0;
              return(
                <tr key={r.id}>
                  <td><span className="badge">{r.code}</span></td>
                  <td style={{fontWeight:600}}>{r.name}</td>
                  <td>{r.class||"—"}</td>
                  <td style={{color:"#22c55e",fontWeight:700}}>{price.toFixed(2)}</td>
                  <td style={{color:C.red}}>{cost.toFixed(4)}</td>
                  <td><span style={{color:margin>30?"#22c55e":margin>15?"#fbbf24":"#f87171",fontWeight:700}}>{margin.toFixed(1)}%</span></td>
                  <td>
                    <button className="btn btn-secondary" style={{fontSize:11,padding:"3px 8px",marginRight:4}} onClick={()=>setViewItem(r)}>{t.view}</button>
                    {hasPerm(mod,"edit")&&<button className="btn btn-secondary" style={{fontSize:11,padding:"3px 8px",marginRight:4}} onClick={()=>doEdit(r)}>{t.edit}</button>}
                    {hasPerm(mod,"delete")&&<button className="btn btn-danger" style={{fontSize:11,padding:"3px 8px"}} onClick={()=>setDelId(r.id)}>{t.delete}</button>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {totalPages>1&&<div style={{display:"flex",justifyContent:"center",gap:6,marginTop:10}}>
        <button className="btn btn-secondary" style={{fontSize:11,padding:"3px 10px"}} disabled={page===1} onClick={()=>setPage(p=>p-1)}>{"<"}</button>
        <span style={{fontSize:12,color:C.muted,padding:"4px 8px"}}>{page} / {totalPages}</span>
        <button className="btn btn-secondary" style={{fontSize:11,padding:"3px 10px"}} disabled={page===totalPages} onClick={()=>setPage(p=>p+1)}>{">"}</button>
      </div>}

      {/* View Modal */}
      {viewItem&&(()=>{
        const cost=calcCost(viewItem.ingredients);
        const price=parseFloat(viewItem.addOnPrice)||0;
        const margin=price>0?((price-cost)/price)*100:0;
        return(
          <div className="overlay" onClick={e=>e.target===e.currentTarget&&setViewItem(null)}>
            <div className="modal modal-lg">
              <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:12}}>{t.view}: {viewItem.name}</h2>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
                <div className="kpi-card"><div className="kpi-label">{t.code}</div><div className="kpi-value">{viewItem.code}</div></div>
                <div className="kpi-card"><div className="kpi-label">{lang==="ar"?"سعر الإضافة":"Add-on Price"}</div><div className="kpi-value" style={{color:"#22c55e"}}>{price.toFixed(2)}</div></div>
                <div className="kpi-card"><div className="kpi-label">{t.cost||"Cost"}</div><div className="kpi-value" style={{color:C.red}}>{cost.toFixed(4)}</div></div>
                <div className="kpi-card"><div className="kpi-label">{lang==="ar"?"الهامش":"Margin"}</div><div className="kpi-value" style={{color:margin>30?"#22c55e":margin>15?"#fbbf24":"#f87171"}}>{margin.toFixed(1)}%</div></div>
                <div className="kpi-card"><div className="kpi-label">{t.class}</div><div className="kpi-value">{viewItem.class||"—"}</div></div>
              </div>
              {viewItem.ingredients&&viewItem.ingredients.length>0&&<>
                <div className="divider"/>
                <div style={{fontWeight:700,fontSize:12,marginBottom:8,color:C.accent}}>{t.ingredients}</div>
                <table><thead><tr><th>{t.code}</th><th>{t.name}</th><th>{t.qty}</th><th>{t.unit}</th><th>{t.cost||"Cost"}</th></tr></thead>
                <tbody>{viewItem.ingredients.map(ing=>{
                  let ingCost=0,ingName="—",ingCode="";
                  const isPrep=ing.source==="prep"||ing.sourceType==="prep";
                  if(!isPrep){const r=rawList.find(x=>String(x.id)===String(ing.srcId)||x.code===ing.code);if(r){const qty=parseFloat(ing.qty)||0;const waste=(parseFloat(ing.waste)||0)/100;const net=qty*(1-waste);ingCost=(r.unit==="piece"?net:net/1000)*parseFloat(r.price||0);ingName=r.name;ingCode=r.code;}}
                  else{const p=prepList.find(x=>String(x.id)===String(ing.srcId)||x.code===ing.code);if(p&&calcPrepCost){const {costPerUnit}=calcPrepCost(p);const qty=parseFloat(ing.qty)||0;const waste=(parseFloat(ing.waste)||0)/100;const net=qty*(1-waste);ingCost=(p.unit==="piece"?net:net/1000)*costPerUnit;ingName=p.name;ingCode=p.code;}}
                  return(<tr key={ing.id}><td>{ingCode||ing.code||"—"}</td><td>{ingName}</td><td>{ing.qty}</td><td>{ing.inputUnit||"g"}</td><td style={{color:C.red}}>{ingCost.toFixed(4)}</td></tr>);
                })}</tbody></table>
              </>}
              <div style={{display:"flex",justifyContent:"flex-end",marginTop:14,gap:8}}>
                {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{setViewItem(null);doEdit(viewItem);}}>{t.edit}</button>}
                <button className="btn btn-secondary" onClick={()=>setViewItem(null)}>{t.cancel}</button>
              </div>
            </div>
          </div>
        );
      })()}

      {showImport&&<ImportModal t={t} lang={lang} C={C} type={t.modifiers||"Modifiers"} onClose={()=>setShowImport(false)} onDownloadTemplate={doDownloadTemplate} onFileSelect={e=>{if(e.target?.files?.[0])doImport(e.target.files[0]);}} showDatePicker={false}/>}

      {showForm&&<div className="overlay" onClick={e=>{if(e.target===e.currentTarget)reset();}}><div className="modal modal-lg" onClick={e=>e.stopPropagation()}>
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.edit:t.add} — {t.modifiers||"Modifiers"}</h2>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
          <div><label className="lbl">{t.code}</label><input value={form.code} onChange={e=>setForm({...form,code:e.target.value})}/></div>
          <div><label className="lbl">{t.name}</label><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">{t.class}</label><select value={form.class} onChange={e=>setForm({...form,class:e.target.value})}><option value="">—</option>{allClasses.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
        </div>
        <div style={{marginBottom:10,background:C.surface,borderRadius:9,padding:"12px 14px",border:`1px solid ${C.border}`}}>
          <label className="lbl">{lang==="ar"?"سعر الإضافة":"Add-on Price"}</label>
          <input type="number" min="0" step="0.01" value={form.addOnPrice} onChange={e=>setForm({...form,addOnPrice:e.target.value})} placeholder="0.00"/>
          {errs.addOnPrice&&<div className="err">{errs.addOnPrice}</div>}
          {parseFloat(form.addOnPrice)>0&&<div style={{fontSize:11,color:live.margin>30?"#22c55e":live.margin>15?"#fbbf24":"#f87171",marginTop:4}}>{lang==="ar"?"هامش:":"Margin:"} {live.margin.toFixed(1)}%</div>}
        </div>
        <div className="divider"/>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <span style={{fontWeight:700,fontSize:13}}>{t.ingredients} <span style={{fontSize:11,color:C.muted}}>({lang==="ar"?"اختياري":"Optional"})</span></span>
          <button className="btn btn-secondary" style={{padding:"5px 12px",fontSize:12}} onClick={addI}>+ {t.addIngredient}</button>
        </div>
        {form.ingredients.map(ing=><IngRowProd key={ing.id} ing={ing} rawList={rawList} prepList={prepList} lang={lang} t={t} C={C} onUpdate={(field,val)=>updI(ing.id,field,val)} onRemove={()=>remI(ing.id)}/>)}
        <div style={{background:C.bg||"#060810",borderRadius:8,padding:"10px 13px",marginTop:8,display:"flex",gap:16,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:C.muted}}>{t.cost||"Cost"}: <strong style={{color:C.red}}>{live.cost.toFixed(4)}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>{lang==="ar"?"الهامش":"Margin"}: <strong style={{color:live.margin>30?"#22c55e":live.margin>15?"#fbbf24":"#f87171"}}>{live.margin.toFixed(1)}%</strong></span>
        </div>
        <div style={{display:"flex",gap:8,marginTop:16}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
      </div></div>}

      {delId&&<DelModal t={t} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}


// ==================== SALES TAB (v22 — full redesign) ====================
const MONTHS=["01","02","03","04","05","06","07","08","09","10","11","12"];

function SalesTab({t,lang,C=DARK,mod,prodList=[],prepList=[],modList=[],rawList=[],hasPerm,salesList,setSalesList,monthlyPrices,setMonthlyPrices,calcProductCost,showToast}){
  const SK_S="tc_sales_v1";
  const SK_MP="tc_monthly_prices_v1";

  // Use props if provided, else self-manage (fallback)
  const [localList,setLocalList]=useState(()=>{try{return JSON.parse(localStorage.getItem(SK_S)||"[]")}catch{return []}});
  const [localMP,setLocalMP]=useState(()=>{try{return JSON.parse(localStorage.getItem(SK_MP)||"{}")}catch{return {}}});
  const list=salesList||localList;
  const mpData=monthlyPrices||localMP;
  const persist=arr=>{
    if(setSalesList)setSalesList(arr); else setLocalList(arr);
    localStorage.setItem(SK_S,JSON.stringify(arr));
  };
  const persistMP=mp=>{
    if(setMonthlyPrices)setMonthlyPrices(mp); else setLocalMP(mp);
    localStorage.setItem(SK_MP,JSON.stringify(mp));
  };

  const [showForm,setShowForm]=useState(false);
  const [editId,setEditId]=useState(null);
  const [delId,setDelId]=useState(null);
  const [showImport,setShowImport]=useState(false);
  const [importDate,setImportDate]=useState({month:String(new Date().getMonth()+1).padStart(2,"0"),year:String(new Date().getFullYear())});
  const [showMPUpload,setShowMPUpload]=useState(false);
  const [mpMonth,setMpMonth]=useState(String(new Date().getMonth()+1).padStart(2,"0"));
  const [mpYear,setMpYear]=useState(String(new Date().getFullYear()));
  const [filterProduct,setFilterProduct]=useState("");
  const [filterType,setFilterType]=useState(""); // "product"|"modifier"|""
  const [filterMonth,setFilterMonth]=useState("");
  const [filterYear,setFilterYear]=useState("");
  const [page,setPage]=useState(1);
  const PAGE_SIZE=25;
  const [selectedIds,setSelectedIds]=useState([]);
  const [showBulkDel,setShowBulkDel]=useState(false);

  // Combined catalog: products + modifiers
  const allItems=useMemo(()=>[
    ...prodList.map(p=>({...p,itemType:"product"})),
    ...(modList||[]).map(m=>({...m,itemType:"modifier",posSellPrice:m.addOnPrice,aggSellPrice:m.addOnPrice}))
  ],[prodList,modList]);

  const blank={
    itemCode:"",itemName:"",itemType:"product",
    qtyPos:"",qtyAgg:"",
    revenuePos:"",revenueAgg:"",
    actualCostPos:"",actualCostAgg:"",
    stdCost:"",costPct:"",
    month:String(new Date().getMonth()+1).padStart(2,"0"),
    year:String(new Date().getFullYear()),
    notes:""
  };
  const [form,setForm]=useState(blank);
  const [errs,setErrs]=useState({});

  // Cost computation
  const getItemCost=useCallback((code,itemType)=>{
    if(itemType==="modifier"){
      const m=(modList||[]).find(x=>x.code===code);
      if(!m||!m.ingredients)return 0;
      return m.ingredients.reduce((s,ing)=>{
        const qtyI=toInternal(parseFloat(ing.qty)||0,ing.inputUnit||"g");
        if(ing.sourceType==="raw"){const r=rawList.find(x=>x.code===ing.code);return s+(r?(qtyI/1000)*parseFloat(r.pricePerKg||0):0);}
        else{const p=prepList.find(x=>x.code===ing.code);if(!p)return s;const pc=p.ingredients?p.ingredients.reduce((ps,pi)=>{const pqI=toInternal(parseFloat(pi.qty)||0,pi.inputUnit||"g");const pr=rawList.find(r=>r.code===pi.code);return ps+(pr?(pqI/1000)*parseFloat(pr.pricePerKg||0):0);},0):0;const yG=parseFloat(p.yieldOverride||p.yield||1000);return s+(yG>0?qtyI*(pc/yG):0);}
      },0);
    }
    if(calcProductCost){
      const p=prodList.find(x=>x.code===code);
      return p?calcProductCost(p):0;
    }
    const prod=prodList.find(p=>p.code===code);
    if(!prod||!prod.ingredients)return 0;
    return prod.ingredients.reduce((s,ing)=>{
      const qtyI=toInternal(parseFloat(ing.qty)||0,ing.inputUnit||"g");
      if(ing.sourceType==="raw"||ing.source==="raw"){const r=rawList.find(x=>x.code===ing.code||String(x.id)===String(ing.rawId||ing.srcId));return s+(r?(qtyI/1000)*parseFloat(r.pricePerKg||0):0);}
      else{const p2=prepList.find(x=>x.code===ing.code||String(x.id)===String(ing.rawId||ing.srcId));if(!p2)return s;const pc=p2.ingredients?p2.ingredients.reduce((ps,pi)=>{const pqI=toInternal(parseFloat(pi.qty)||0,pi.inputUnit||"g");const pr=rawList.find(r=>r.code===pi.code||String(r.id)===String(pi.rawId||pi.srcId));return ps+(pr?(pqI/1000)*parseFloat(pr.pricePerKg||0):0);},0):0;const yG=parseFloat(p2.yieldOverride||p2.yield||1000);return s+(yG>0?qtyI*(pc/yG):0);}
    },0);
  },[prodList,modList,rawList,prepList,calcProductCost]);

  const getStdCost=useCallback((code,itemType)=>{
    if(itemType==="modifier")return 0;
    const prod=prodList.find(p=>p.code===code);
    return parseFloat(prod?.stdCost||0);
  },[prodList]);

  // Live totals for form
  const liveCost=useMemo(()=>{
    const unitCost=getItemCost(form.itemCode,form.itemType);
    const qPos=parseFloat(form.qtyPos)||0;
    const qAgg=parseFloat(form.qtyAgg)||0;
    const rPos=parseFloat(form.revenuePos)||0;
    const rAgg=parseFloat(form.revenueAgg)||0;
    const cPos=unitCost*qPos;
    const cAgg=unitCost*qAgg;
    const totalRev=rPos+rAgg;
    const totalCost=cPos+cAgg;
    const costPct=totalRev>0?(totalCost/totalRev)*100:0;
    return {unitCost,cPos,cAgg,totalRev,totalCost,costPct};
  },[form,getItemCost]);

  const handleItemSelect=useCallback((code)=>{
    if(!code){setForm(f=>({...f,itemCode:"",itemName:"",itemType:"product"}));return;}
    const item=allItems.find(x=>x.code===code);
    if(!item)return;
    const posP=parseFloat(item.posSellPrice||item.addOnPrice||0);
    const aggP=parseFloat(item.aggSellPrice||item.addOnPrice||0);
    setForm(f=>{
      const qPos=parseFloat(f.qtyPos)||0;
      const qAgg=parseFloat(f.qtyAgg)||0;
      return {...f,
        itemCode:code,itemName:item.name,itemType:item.itemType,
        revenuePos:posP>0&&qPos>0?String((posP*qPos).toFixed(2)):posP>0?String(posP.toFixed(2)):f.revenuePos,
        revenueAgg:aggP>0&&qAgg>0?String((aggP*qAgg).toFixed(2)):aggP>0?String(aggP.toFixed(2)):f.revenueAgg,
      };
    });
  },[allItems]);

  const validate=()=>{
    const e={};
    if(!form.itemCode&&!form.itemName)e.item=lang==="ar"?"اختر منتجاً":"Select an item";
    const hasQty=(parseFloat(form.qtyPos)||0)+(parseFloat(form.qtyAgg)||0);
    if(!hasQty)e.qty=lang==="ar"?"أدخل كمية في قناة واحدة على الأقل":"Enter qty for at least one channel";
    setErrs(e);
    return !Object.keys(e).length;
  };

  const save=()=>{
    if(!validate())return;
    const unitCost=getItemCost(form.itemCode,form.itemType);
    const qPos=parseFloat(form.qtyPos)||0;
    const qAgg=parseFloat(form.qtyAgg)||0;
    const rec={
      ...form,
      id:editId||Date.now(),
      actualCostPos:(unitCost*qPos).toFixed(4),
      actualCostAgg:(unitCost*qAgg).toFixed(4),
      stdCost:getStdCost(form.itemCode,form.itemType).toFixed(4),
    };
    if(editId)persist(list.map(r=>r.id===editId?rec:r));
    else persist([...list,rec]);
    showToast&&showToast(lang==="ar"?"تم الحفظ":"Saved");
    reset();
  };

  const reset=()=>{setForm(blank);setEditId(null);setShowForm(false);setErrs({});};
  const doEdit=item=>{setForm({...item});setEditId(item.id);setShowForm(true);};
  const doDelete=id=>{persist(list.filter(r=>r.id!==id));setDelId(null);showToast&&showToast(lang==="ar"?"تم الحذف":"Deleted","red");};
  const doBulkDelete=()=>{
    persist(list.filter(r=>!selectedIds.includes(r.id)));
    showToast&&showToast(`${lang==="ar"?"تم حذف":"Deleted"} ${selectedIds.length}`,"red");
    setSelectedIds([]);
    setShowBulkDel(false);
  };

  const handleMPFile=file=>{
    const reader=new FileReader();
    reader.onload=e=>{
      const wb=XLSX.read(e.target.result,{type:"binary"});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      const key=`${mpYear}-${mpMonth}`;
      const prices={};
      rows.forEach(row=>{const code=(row.code||row.Code||"").toString().trim();const price=parseFloat(row.pricePerKg||row.price_per_kg||row.price||0);if(code)prices[code]=price;});
      const updated={...mpData,[key]:prices};
      persistMP(updated);
      showToast&&showToast(`${lang==="ar"?"تم حفظ أسعار":"Prices saved:"} ${key} (${Object.keys(prices).length})`);
      setShowMPUpload(false);
    };
    reader.readAsBinaryString(file);
  };

  const doDownloadTemplate=()=>{
    const headers=["itemCode","itemName","itemType","qtyPos","qtyAgg","revenuePos","revenueAgg","month","year","notes"];
    const example=["P001","Chicken Burger","product","50","20","2500","1200","06","2026",""];
    const ws=XLSX.utils.aoa_to_sheet([headers,example]);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Sales");
    XLSX.writeFile(wb,"sales_template.xlsx");
  };

  const numVal=v=>{
    if(v===undefined||v===null||v==="")return 0;
    const n=Number(String(v).replace(/,/g,"").trim());
    return isNaN(n)?0:n;
  };

  const doImport=file=>{
    if(!file){showToast&&showToast(lang==="ar"?"لم يتم اختيار ملف":"No file selected","warning");return;}
    const reader=new FileReader();
    reader.onerror=()=>{showToast&&showToast(lang==="ar"?"فشل قراءة الملف":"Failed to read file","warning");};
    reader.onload=e=>{
     try{
      const wb=XLSX.read(e.target.result,{type:"binary"});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      if(!rows.length){showToast&&showToast(lang==="ar"?"الملف فارغ أو بدون بيانات صالحة":"File is empty or has no valid rows","warning");return;}
      let added=0,updatedCount=0;
      const updated=[...list];
      const defMonth=importDate?.month||String(new Date().getMonth()+1).padStart(2,"0");
      const defYear=importDate?.year||String(new Date().getFullYear());
      // Find an existing row for the same item + period, so re-importing UPDATES instead of duplicating
      const findExistingIdx=(code,name,type,m,y)=>updated.findIndex(x=>x.month===m&&x.year===y&&(x.itemType||"product")===type&&((code&&x.itemCode===code)||(!code&&x.itemName===name)));
      rows.forEach(row=>{
        // Support old format: productCode/channel/qty/revenue
        const isOldFmt=!!(row.channel||row.Channel);
        let rec,code,name,type,m,y;
        if(isOldFmt){
          code=(row.productCode||row.product_code||row.itemCode||"").toString().trim();
          name=(row.productName||row.product_name||row.itemName||"").toString().trim();
          if(!code&&!name)return;
          type="product";
          const ch=(row.channel||"POS").toString().toUpperCase();
          const qty=numVal(row.qty)||1;
          m=(row.month||defMonth).toString().padStart(2,"0");
          y=(row.year||defYear).toString();
          const item=allItems.find(x=>x.code===code||x.name===name);
          const unitCost=item?numVal(getItemCost(item.code,item.itemType)):0;
          const price=ch==="POS"?numVal(item?.posSellPrice):numVal(item?.aggSellPrice);
          const rev=numVal(row.revenue)||(price*qty);
          rec={itemCode:code,itemName:name,itemType:"product",
            qtyPos:ch==="POS"?qty:0,qtyAgg:ch==="AGG"?qty:0,
            revenuePos:ch==="POS"?rev.toFixed(2):"0",revenueAgg:ch==="AGG"?rev.toFixed(2):"0",
            actualCostPos:ch==="POS"?(unitCost*qty).toFixed(4):"0",actualCostAgg:ch==="AGG"?(unitCost*qty).toFixed(4):"0",
            stdCost:"0",costPct:"0",month:m,year:y,notes:row.notes||""};
        } else {
          code=(row.itemCode||row.productCode||"").toString().trim();
          name=(row.itemName||row.productName||"").toString().trim();
          if(!code&&!name)return;
          type=(row.itemType||"product").toString().toLowerCase();
          const qPos=numVal(row.qtyPos);
          const qAgg=numVal(row.qtyAgg);
          m=(row.month||defMonth).toString().padStart(2,"0");
          y=(row.year||defYear).toString();
          const item=allItems.find(x=>x.code===code||x.name===name);
          // If revenue columns are blank, auto-calculate from the item's POS/AGG selling price (Products/Modifiers modules)
          const rPos=row.revenuePos!==undefined&&row.revenuePos!==""?numVal(row.revenuePos):numVal(item?.posSellPrice)*qPos;
          const rAgg=row.revenueAgg!==undefined&&row.revenueAgg!==""?numVal(row.revenueAgg):numVal(item?.aggSellPrice)*qAgg;
          const unitCost=item?numVal(getItemCost(item.code||code,item.itemType||type)):0;
          const stdCostVal=item?numVal(getStdCost(item.code||code,item.itemType||type)):0;
          rec={itemCode:code,itemName:name,itemType:type,
            qtyPos:qPos,qtyAgg:qAgg,revenuePos:rPos.toFixed(2),revenueAgg:rAgg.toFixed(2),
            actualCostPos:(unitCost*qPos).toFixed(4),actualCostAgg:(unitCost*qAgg).toFixed(4),
            stdCost:stdCostVal.toFixed(4),
            costPct:"0",month:m,year:y,notes:row.notes||""};
        }
        const existingIdx=findExistingIdx(code,name,type,m,y);
        if(existingIdx>=0){
          updated[existingIdx]={...updated[existingIdx],...rec,id:updated[existingIdx].id};
          updatedCount++;
        } else {
          updated.push({...rec,id:Date.now()+Math.random()});
          added++;
        }
      });
      persist(updated);
      showToast&&showToast(`${lang==="ar"?"مُضاف:":"Added:"} ${added} | ${lang==="ar"?"مُحدّث:":"Updated:"} ${updatedCount}`);
      setShowImport(false);
     }catch(err){
      console.error(err);
      showToast&&showToast(lang==="ar"?"تعذر استيراد الملف - تأكد أنه بصيغة Excel صحيحة":"Could not import file - make sure it's a valid Excel file","warning");
     }
    };
    reader.readAsBinaryString(file);
  };

  const doExport=()=>{
    const rows=filtered.map(r=>{
      const rPos=parseFloat(r.revenuePos||0);
      const rAgg=parseFloat(r.revenueAgg||0);
      const cPos=parseFloat(r.actualCostPos||0);
      const cAgg=parseFloat(r.actualCostAgg||0);
      const totalRev=rPos+rAgg;
      const totalCost=cPos+cAgg;
      return{itemCode:r.itemCode,itemName:r.itemName,itemType:r.itemType,
        qtyPos:r.qtyPos,qtyAgg:r.qtyAgg,
        revenuePos:rPos.toFixed(2),revenueAgg:rAgg.toFixed(2),totalRevenue:totalRev.toFixed(2),
        actualCostPos:cPos.toFixed(4),actualCostAgg:cAgg.toFixed(4),totalActualCost:totalCost.toFixed(4),
        stdCost:r.stdCost,costPct:totalRev>0?((totalCost/totalRev)*100).toFixed(1)+"%":"0%",
        month:r.month,year:r.year,notes:r.notes||""};
    });
    const ws=XLSX.utils.json_to_sheet(rows);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Sales");
    XLSX.writeFile(wb,"sales_export.xlsx");
  };

  const filtered=useMemo(()=>list.filter(r=>{
    if(filterType&&r.itemType!==filterType)return false;
    if(filterProduct){const q=filterProduct.toLowerCase();if(!(r.itemName||"").toLowerCase().includes(q)&&!(r.itemCode||"").includes(q))return false;}
    if(filterMonth&&r.month!==filterMonth)return false;
    if(filterYear&&r.year!==filterYear)return false;
    return true;
  }),[list,filterType,filterProduct,filterMonth,filterYear]);

  // KPI summary
  const kpi=useMemo(()=>{
    let rPosTot=0,rAggTot=0,cPosTot=0,cAggTot=0;
    filtered.forEach(r=>{
      rPosTot+=parseFloat(r.revenuePos||0);
      rAggTot+=parseFloat(r.revenueAgg||0);
      cPosTot+=parseFloat(r.actualCostPos||0);
      cAggTot+=parseFloat(r.actualCostAgg||0);
    });
    const totalRev=rPosTot+rAggTot;
    const totalCost=cPosTot+cAggTot;
    const costPct=totalRev>0?(totalCost/totalRev)*100:0;
    const gp=totalRev-totalCost;
    return {rPosTot,rAggTot,cPosTot,cAggTot,totalRev,totalCost,costPct,gp};
  },[filtered]);

  const totalPages=Math.max(1,Math.ceil(filtered.length/PAGE_SIZE));
  const paged=filtered.slice((page-1)*PAGE_SIZE,page*PAGE_SIZE);

  const ar=lang==="ar";

  return(
    <div>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:8}}>
        <h2 style={{fontSize:15,fontWeight:700,color:C.accent,margin:0}}>{t.sales||"Sales"} <span style={{fontSize:12,color:C.muted}}>({list.length})</span></h2>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <button className="btn btn-secondary" style={{fontSize:12,background:"#7c3aed22",borderColor:"#7c3aed",color:"#a78bfa"}} onClick={()=>setShowMPUpload(true)}>
            {ar?"المبيعات الشهرية":"Monthly Sales"}
          </button>
          {selectedIds.length>0&&hasPerm(mod,"delete")&&<button className="btn btn-danger" style={{fontSize:12}} onClick={()=>setShowBulkDel(true)}>{ar?`حذف المحدد (${selectedIds.length})`:`Delete Selected (${selectedIds.length})`}</button>}
          {hasPerm(mod,"import")&&<button className="btn btn-secondary" style={{fontSize:12}} onClick={()=>setShowImport(true)}>{t.import}</button>}
          {hasPerm(mod,"export")&&<button className="btn btn-secondary" style={{fontSize:12}} onClick={doExport}>{t.export}</button>}
          {hasPerm(mod,"add")&&<button className="btn btn-primary" style={{fontSize:12}} onClick={()=>setShowForm(true)}>+ {t.add}</button>}
        </div>
      </div>

      {/* KPI Cards — 4 cards, POS/AGG split below each */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
        {/* Total Sales */}
        <div className="kpi-card">
          <div className="kpi-label">{ar?"إجمالي المبيعات":"Total Sales"}</div>
          <div className="kpi-value" style={{color:"#22c55e"}}>{kpi.totalRev.toFixed(2)}</div>
          <div style={{display:"flex",justifyContent:"space-between",marginTop:6,paddingTop:6,borderTop:`1px solid ${C.border}`}}>
            <span style={{fontSize:10,color:"#22c55e"}}>POS: {kpi.rPosTot.toFixed(2)}</span>
            <span style={{fontSize:10,color:"#60a5fa"}}>AGG: {kpi.rAggTot.toFixed(2)}</span>
          </div>
        </div>
        {/* Total Cost */}
        <div className="kpi-card">
          <div className="kpi-label">{ar?"إجمالي التكلفة":"Total Cost"}</div>
          <div className="kpi-value" style={{color:C.red}}>{kpi.totalCost.toFixed(2)}</div>
          <div style={{display:"flex",justifyContent:"space-between",marginTop:6,paddingTop:6,borderTop:`1px solid ${C.border}`}}>
            <span style={{fontSize:10,color:"#f87171"}}>POS: {kpi.cPosTot.toFixed(2)}</span>
            <span style={{fontSize:10,color:"#f87171"}}>AGG: {kpi.cAggTot.toFixed(2)}</span>
          </div>
        </div>
        {/* Cost % */}
        <div className="kpi-card">
          <div className="kpi-label">{ar?"نسبة التكلفة":"Cost %"}</div>
          <div className="kpi-value" style={{color:kpi.costPct<30?"#22c55e":kpi.costPct<40?"#fbbf24":"#f87171"}}>{kpi.costPct.toFixed(1)}%</div>
          <div style={{marginTop:6,paddingTop:6,borderTop:`1px solid ${C.border}`,fontSize:10,color:C.muted}}>
            {ar?"من إجمالي الإيرادات":"Of total revenue"}
          </div>
        </div>
        {/* Gross Profit / Loss */}
        <div className="kpi-card">
          <div className="kpi-label">{ar?"مجمل الربح":"Gross Profit"}</div>
          <div className="kpi-value" style={{color:kpi.gp>=0?"#22c55e":"#f87171"}}>{Math.abs(kpi.gp).toFixed(2)}</div>
          <div style={{marginTop:6,paddingTop:6,borderTop:`1px solid ${C.border}`,fontSize:10,color:kpi.gp>=0?"#22c55e":"#f87171",fontWeight:700}}>
            {kpi.gp>=0?(ar?"ربح":"Profit"):(ar?"خسارة":"Loss")} · {(100-kpi.costPct).toFixed(1)}%
          </div>
        </div>
      </div>

      {/* Filters */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:10,padding:"10px 14px",marginBottom:12}}>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <input placeholder={ar?"بحث منتج/كود":"Search item/code"} value={filterProduct} onChange={e=>{setFilterProduct(e.target.value);setPage(1);}} style={{flex:1,minWidth:140,maxWidth:240}}/>
          <div style={{display:"flex",gap:4,alignItems:"center",background:C.surface,borderRadius:7,padding:"3px 5px",border:`1px solid ${C.border}`}}>
            {[{v:"",l:ar?"الكل":"All"},{v:"product",l:ar?"منتجات":"Products"},{v:"modifier",l:ar?"مودفاير":"Modifiers"}].map(opt=>(
              <button key={opt.v} onClick={()=>{setFilterType(opt.v);setPage(1);}} style={{background:filterType===opt.v?C.accent:"transparent",color:filterType===opt.v?"#080b14":C.muted,border:"none",borderRadius:5,padding:"4px 12px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit",transition:"all .15s"}}>{opt.l}</button>
            ))}
          </div>
          <select value={filterMonth} onChange={e=>{setFilterMonth(e.target.value);setPage(1);}} style={{minWidth:110,maxWidth:140}}>
            <option value="">{ar?"كل الشهور":"All Months"}</option>
            {MONTHS.map(m=><option key={m} value={m}>{mLabel(m,lang)}</option>)}
          </select>
          <select value={filterYear} onChange={e=>{setFilterYear(e.target.value);setPage(1);}} style={{minWidth:80,maxWidth:100}}>
            <option value="">{ar?"كل السنوات":"All Years"}</option>
            {["2024","2025","2026","2027"].map(y=><option key={y} value={y}>{y}</option>)}
          </select>
          {(filterType||filterProduct||filterMonth||filterYear)&&<button className="btn btn-secondary" style={{fontSize:11,padding:"5px 10px"}} onClick={()=>{setFilterType("");setFilterProduct("");setFilterMonth("");setFilterYear("");setPage(1);}}>{ar?"مسح":"Clear"}</button>}
        </div>
      </div>

      {/* Table */}
      <div className="table-wrap">
        <table>
          <thead><tr>
            <th style={{width:32}}><input type="checkbox" checked={paged.length>0&&paged.every(r=>selectedIds.includes(r.id))} onChange={e=>{
              if(e.target.checked)setSelectedIds(prev=>[...new Set([...prev,...paged.map(r=>r.id)])]);
              else setSelectedIds(prev=>prev.filter(id=>!paged.some(r=>r.id===id)));
            }}/></th>
            <th>{t.productName}</th>
            <th>{ar?"النوع":"Type"}</th>
            <th>POS Qty</th>
            <th>AGG Qty</th>
            <th>POS {ar?"مبيعات":"Rev."}</th>
            <th>AGG {ar?"مبيعات":"Rev."}</th>
            <th>POS {ar?"تكلفة":"Cost"}</th>
            <th>AGG {ar?"تكلفة":"Cost"}</th>
            <th>{ar?"تكلفة معيارية":"Std Cost"}</th>
            <th>{ar?"نسبة التكلفة":"Cost %"}</th>
            <th>{ar?"الشهر":"Month"}</th>
            <th>{t.actions}</th>
          </tr></thead>
          <tbody>
            {paged.length===0&&<tr><td colSpan={13} style={{textAlign:"center",color:C.muted,padding:20}}>{t.noData}</td></tr>}
            {paged.map(r=>{
              const rPos=parseFloat(r.revenuePos||0);
              const rAgg=parseFloat(r.revenueAgg||0);
              const cPos=parseFloat(r.actualCostPos||0);
              const cAgg=parseFloat(r.actualCostAgg||0);
              const totalRev=rPos+rAgg;
              const totalCost=cPos+cAgg;
              const costPct=totalRev>0?(totalCost/totalRev)*100:0;
              return(
                <tr key={r.id}>
                  <td><input type="checkbox" checked={selectedIds.includes(r.id)} onChange={e=>{
                    if(e.target.checked)setSelectedIds(prev=>[...prev,r.id]);
                    else setSelectedIds(prev=>prev.filter(id=>id!==r.id));
                  }}/></td>
                  <td style={{fontWeight:600}}>{r.itemName||r.productName||r.itemCode}</td>
                  <td><span style={{background:r.itemType==="modifier"?"#7c3aed22":"#22c55e22",color:r.itemType==="modifier"?"#a78bfa":"#22c55e",padding:"2px 8px",borderRadius:5,fontSize:10,fontWeight:700}}>{r.itemType==="modifier"?(ar?"مودفاير":"Modifier"):(ar?"منتج":"Product")}</span></td>
                  <td style={{color:C.muted}}>{r.qtyPos||0}</td>
                  <td style={{color:C.muted}}>{r.qtyAgg||0}</td>
                  <td style={{color:"#22c55e",fontWeight:600}}>{rPos.toFixed(2)}</td>
                  <td style={{color:"#60a5fa",fontWeight:600}}>{rAgg.toFixed(2)}</td>
                  <td style={{color:C.red,fontSize:11}}>{cPos.toFixed(4)}</td>
                  <td style={{color:C.red,fontSize:11}}>{cAgg.toFixed(4)}</td>
                  <td style={{color:C.muted,fontSize:11}}>{parseFloat(r.stdCost||0).toFixed(4)}</td>
                  <td><span style={{color:costPct<30?"#22c55e":costPct<40?"#fbbf24":"#f87171",fontWeight:700}}>{costPct.toFixed(1)}%</span></td>
                  <td style={{color:C.muted,fontSize:11}}>{mLabel(r.month,lang)} {r.year}</td>
                  <td>
                    {hasPerm(mod,"edit")&&<button className="btn btn-secondary" style={{fontSize:11,padding:"3px 8px",marginRight:4}} onClick={()=>doEdit(r)}>{t.edit}</button>}
                    {hasPerm(mod,"delete")&&<button className="btn btn-danger" style={{fontSize:11,padding:"3px 8px"}} onClick={()=>setDelId(r.id)}>{t.delete}</button>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {totalPages>1&&<div style={{display:"flex",justifyContent:"center",gap:6,marginTop:10}}>
        <button className="btn btn-secondary" style={{fontSize:11,padding:"3px 10px"}} disabled={page===1} onClick={()=>setPage(p=>p-1)}>{"<"}</button>
        <span style={{fontSize:12,color:C.muted,padding:"4px 8px"}}>{page} / {totalPages}</span>
        <button className="btn btn-secondary" style={{fontSize:11,padding:"3px 10px"}} disabled={page===totalPages} onClick={()=>setPage(p=>p+1)}>{">"}</button>
      </div>}

      {/* Monthly Prices / Sales Upload Modal */}
      {showMPUpload&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&setShowMPUpload(false)}>
        <div className="modal">
          <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:14}}>{ar?"المبيعات الشهرية":"Monthly Sales"}</h2>
          <div style={{marginBottom:12}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
              <div><label className="lbl">{ar?"الشهر":"Month"}</label>
                <select value={mpMonth} onChange={e=>setMpMonth(e.target.value)}>
                  {MONTHS.map(m=><option key={m} value={m}>{mLabel(m,lang)}</option>)}
                </select>
              </div>
              <div><label className="lbl">{ar?"السنة":"Year"}</label>
                <select value={mpYear} onChange={e=>setMpYear(e.target.value)}>
                  {["2024","2025","2026","2027"].map(y=><option key={y} value={y}>{y}</option>)}
                </select>
              </div>
            </div>
            <div style={{fontSize:12,color:C.muted,marginBottom:10}}>
              {ar?"الملف يجب أن يحتوي على: code, pricePerKg":"File must contain: code, pricePerKg"}
            </div>
            {Object.keys(mpData).length>0&&<div style={{marginBottom:10}}>
              <div style={{fontSize:11,color:C.accent,marginBottom:4}}>{ar?"البيانات المحفوظة:":"Saved periods:"}</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                {Object.keys(mpData).sort().map(k=>(
                  <span key={k} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:6,padding:"2px 8px",fontSize:11}}>{k} ({Object.keys(mpData[k]).length})</span>
                ))}
              </div>
            </div>}
            <div style={{background:C.surface,border:`2px dashed ${C.border}`,borderRadius:9,padding:"16px",textAlign:"center"}}>
              <label style={{cursor:"pointer",color:C.accent,fontSize:13}}>
                {ar?"اختر ملف Excel":"Choose Excel File"}
                <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{if(e.target.files[0])handleMPFile(e.target.files[0]);}}/>
              </label>
            </div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:12}}>
            <button className="btn btn-secondary" style={{flex:1}} onClick={()=>setShowMPUpload(false)}>{t.cancel}</button>
          </div>
        </div>
      </div>}

      {showImport&&<ImportModal t={t} lang={lang} C={C} type={t.sales||"Sales"} onClose={()=>setShowImport(false)} onDownloadTemplate={doDownloadTemplate} onFileSelect={e=>{if(e.target?.files?.[0])doImport(e.target.files[0]);}} showDatePicker={true} onDateSelect={setImportDate}/>}

      {/* Add/Edit Form Modal */}
      {showForm&&<div className="overlay" onClick={e=>{if(e.target===e.currentTarget)reset();}}><div className="modal modal-lg" onClick={e=>e.stopPropagation()}>
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.edit:t.add} — {t.sales||"Sales"}</h2>

        {/* Item selection */}
        <div style={{marginBottom:10}}>
          <label className="lbl">{ar?"المنتج / المودفاير":"Product / Modifier"}</label>
          <select value={form.itemCode} onChange={e=>handleItemSelect(e.target.value)}>
            <option value="">— {ar?"اختر عنصراً":"Select item"} —</option>
            {prodList.length>0&&<optgroup label={ar?"المنتجات":"Products"}>{prodList.map(p=><option key={p.id} value={p.code}>{p.name} ({p.code})</option>)}</optgroup>}
            {(modList||[]).length>0&&<optgroup label={ar?"المودفاير":"Modifiers"}>{(modList||[]).map(m=><option key={m.id} value={m.code}>{m.name} ({m.code})</option>)}</optgroup>}
          </select>
          {errs.item&&<div className="err">{errs.item}</div>}
        </div>

        {/* Qty + Revenue side by side POS/AGG */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10,background:C.surface,borderRadius:9,padding:"12px 14px",border:`1px solid ${C.border}`}}>
          <div>
            <label className="lbl" style={{color:"#22c55e"}}>POS — {ar?"الكمية":"Qty"}</label>
            <input type="number" min="0" step="1" value={form.qtyPos} onChange={e=>{
              const q=parseFloat(e.target.value)||0;
              const item=allItems.find(x=>x.code===form.itemCode);
              const p=parseFloat(item?.posSellPrice||item?.addOnPrice||0);
              setForm(f=>({...f,qtyPos:e.target.value,revenuePos:p>0?(p*q).toFixed(2):f.revenuePos}));
            }} placeholder="0"/>
          </div>
          <div>
            <label className="lbl" style={{color:"#22c55e"}}>POS — {ar?"المبيعات":"Revenue"}</label>
            <input type="number" min="0" step="0.01" value={form.revenuePos} onChange={e=>setForm(f=>({...f,revenuePos:e.target.value}))} placeholder="0.00"/>
          </div>
          <div>
            <label className="lbl" style={{color:"#60a5fa"}}>AGG — {ar?"الكمية":"Qty"}</label>
            <input type="number" min="0" step="1" value={form.qtyAgg} onChange={e=>{
              const q=parseFloat(e.target.value)||0;
              const item=allItems.find(x=>x.code===form.itemCode);
              const p=parseFloat(item?.aggSellPrice||item?.addOnPrice||0);
              setForm(f=>({...f,qtyAgg:e.target.value,revenueAgg:p>0?(p*q).toFixed(2):f.revenueAgg}));
            }} placeholder="0"/>
          </div>
          <div>
            <label className="lbl" style={{color:"#60a5fa"}}>AGG — {ar?"المبيعات":"Revenue"}</label>
            <input type="number" min="0" step="0.01" value={form.revenueAgg} onChange={e=>setForm(f=>({...f,revenueAgg:e.target.value}))} placeholder="0.00"/>
          </div>
        </div>
        {errs.qty&&<div className="err" style={{marginTop:-6,marginBottom:8}}>{errs.qty}</div>}

        {/* Month/Year */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
          <div><label className="lbl">{ar?"الشهر":"Month"}</label>
            <select value={form.month} onChange={e=>setForm(f=>({...f,month:e.target.value}))}>
              {MONTHS.map(m=><option key={m} value={m}>{mLabel(m,lang)}</option>)}
            </select>
          </div>
          <div><label className="lbl">{ar?"السنة":"Year"}</label>
            <select value={form.year} onChange={e=>setForm(f=>({...f,year:e.target.value}))}>
              {["2024","2025","2026","2027"].map(y=><option key={y} value={y}>{y}</option>)}
            </select>
          </div>
          <div><label className="lbl">{ar?"ملاحظات":"Notes"} ({ar?"اختياري":"Optional"})</label>
            <input value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))}/>
          </div>
        </div>

        {/* Live cost preview */}
        {form.itemCode&&<div style={{background:C.bg||"#060810",borderRadius:8,padding:"10px 13px",marginTop:4,display:"flex",gap:16,flexWrap:"wrap"}}>
          <span style={{fontSize:12,color:C.muted}}>{ar?"تكلفة الوحدة":"Unit Cost"}: <strong style={{color:C.red}}>{liveCost.unitCost.toFixed(4)}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>POS {ar?"تكلفة":"Cost"}: <strong style={{color:C.red}}>{liveCost.cPos.toFixed(4)}</strong></span>
          <span style={{fontSize:12,color:C.muted}}>AGG {ar?"تكلفة":"Cost"}: <strong style={{color:C.red}}>{liveCost.cAgg.toFixed(4)}</strong></span>
          {liveCost.totalRev>0&&<span style={{fontSize:12,color:C.muted}}>{ar?"نسبة التكلفة":"Cost %"}: <strong style={{color:liveCost.costPct<30?"#22c55e":liveCost.costPct<40?"#fbbf24":"#f87171"}}>{liveCost.costPct.toFixed(1)}%</strong></span>}
        </div>}

        <div style={{display:"flex",gap:8,marginTop:16}}>
          <button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button>
          <button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button>
        </div>
      </div></div>}

      {delId&&<DelModal t={t} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
      {showBulkDel&&<div className="overlay"><div className="modal" style={{maxWidth:320,textAlign:"center"}}>
        <p style={{marginBottom:20,color:C.text,fontSize:14}}>{ar?`هل تريد حذف ${selectedIds.length} سجل مبيعات؟`:`Delete ${selectedIds.length} sales records?`}</p>
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          <button className="btn-sm-d" style={{padding:"8px 20px",fontSize:13}} onClick={doBulkDelete}>{t.delete}</button>
          <button className="btn btn-secondary" style={{padding:"8px 20px",fontSize:13}} onClick={()=>setShowBulkDel(false)}>{t.cancel}</button>
        </div>
      </div></div>}
    </div>
  );
}


// ==================== CLASSES TAB ====================
function ClassesTab({t,lang,C=DARK,mod,classes:clsObj,setClasses:setClsObj,hasPerm}){
  const [showForm,setShowForm]=useState(false);
  const [newClass,setNewClass]=useState("");
  const [activeSection,setActiveSection]=useState("products");
  const [editIdx,setEditIdx]=useState(null);
  const [err,setErr]=useState("");
  const ar=lang==="ar";

  const sections=[
    {key:"products",label:ar?"منتجات":"Products"},
    {key:"prep",label:ar?"شبه مصنع":"Prep"},
    {key:"raw",label:ar?"مواد خام":"Raw"},
    {key:"modifiers",label:ar?"موديفايرز":"Modifiers"},
  ];

  const curList=clsObj[activeSection]||[];
  const persist=(sec,arr)=>{const updated={...clsObj,[sec]:arr};setClsObj(updated);localStorage.setItem("tc_cls_v1",JSON.stringify(updated));};

  const save=()=>{
    const v=newClass.trim();
    if(!v){setErr(ar?"مطلوب":"Required");return;}
    if(editIdx!==null){
      const updated=[...curList];updated[editIdx]=v;persist(activeSection,updated);
    } else {
      if(curList.includes(v)){setErr(ar?"مكرر":"Duplicate");return;}
      persist(activeSection,[...curList,v]);
    }
    setNewClass("");setShowForm(false);setEditIdx(null);setErr("");
  };

  const doDelete=idx=>persist(activeSection,curList.filter((_,i)=>i!==idx));
  const doEdit=idx=>{setNewClass(curList[idx]);setEditIdx(idx);setShowForm(true);};
  const totalCount=Object.values(clsObj).reduce((s,a)=>s+(Array.isArray(a)?a.length:0),0);

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <h2 style={{fontSize:15,fontWeight:700,color:C.accent,margin:0}}>{t.classes} <span style={{fontSize:12,color:C.muted}}>({totalCount})</span></h2>
        {hasPerm(mod,"add")&&<button className="btn btn-primary" style={{fontSize:12}} onClick={()=>{setNewClass("");setEditIdx(null);setShowForm(true);}}>+ {ar?"إضافة فئة":"Add Class"}</button>}
      </div>
      <div style={{display:"flex",gap:6,marginBottom:14}}>
        {sections.map(s=>(
          <button key={s.key} onClick={()=>{setActiveSection(s.key);setShowForm(false);setErr("");}}
            style={{background:activeSection===s.key?C.accent:"transparent",color:activeSection===s.key?"#080b14":C.muted,border:`1px solid ${activeSection===s.key?C.accent:C.border}`,borderRadius:8,padding:"6px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
            {s.label} ({(clsObj[s.key]||[]).length})
          </button>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10}}>
        {curList.length===0&&<div style={{color:C.muted,fontSize:13,padding:"20px 0"}}>{t.noData}</div>}
        {curList.map((c,i)=>(
          <div key={i} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:10,padding:"12px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontWeight:600,fontSize:13}}>{c}</span>
            <div style={{display:"flex",gap:4}}>
              {hasPerm(mod,"edit")&&<button className="btn btn-secondary" style={{fontSize:11,padding:"3px 8px"}} onClick={()=>doEdit(i)}>{t.edit}</button>}
              {hasPerm(mod,"delete")&&<button className="btn btn-danger" style={{fontSize:11,padding:"3px 8px"}} onClick={()=>doDelete(i)}>{t.delete}</button>}
            </div>
          </div>
        ))}
      </div>
      {showForm&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&setShowForm(false)}>
        <div className="modal">
          <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:14}}>{editIdx!==null?t.edit:ar?"إضافة فئة":"Add Class"} — {sections.find(s=>s.key===activeSection)?.label}</h2>
          <input value={newClass} onChange={e=>{setNewClass(e.target.value);setErr("");}} placeholder={ar?"اسم الفئة":"Class name"} autoFocus/>
          {err&&<div className="err">{err}</div>}
          <div style={{display:"flex",gap:8,marginTop:14}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={()=>{setShowForm(false);setErr("");}}>{t.cancel}</button></div>
        </div>
      </div>}
    </div>
  );
}

// ==================== USERS TAB ====================
const ALL_PERMS=["raw","prep","products","modifiers","sales","classes"];
function UsersTab({t,lang,C=DARK,users:usersList,setUsers:setUsersList,currentUserId}){
  const mod="users";
  const currentUser=usersList.find(u=>u.id===currentUserId);
  const hasPerm=(_m,_a)=>true;
  const [showForm,setShowForm]=useState(false);
  const [editId,setEditId]=useState(null);
  const [delId,setDelId]=useState(null);
  const blank={id:"",pin:"",name:"",role:"viewer",permissions:{}};
  const [form,setForm]=useState(blank);
  const [errs,setErrs]=useState({});

  const persist=arr=>{setUsersList(arr);localStorage.setItem("tc_users_v1",JSON.stringify(arr));};

  const DEFAULT_PERMS=(role)=>{
    if(role==="admin")return Object.fromEntries(ALL_PERMS.map(k=>[k,{view:true,add:true,edit:true,delete:true,import:true,export:true}]));
    if(role==="manager")return Object.fromEntries(ALL_PERMS.map(k=>[k,{view:true,add:true,edit:true,delete:false,import:true,export:true}]));
    return Object.fromEntries(ALL_PERMS.map(k=>[k,{view:true,add:false,edit:false,delete:false,import:false,export:true}]));
  };

  const validate=()=>{
    const e={};
    if(!form.id.toString().trim())e.id=lang==="ar"?"مطلوب":"Required";
    if(!form.name.trim())e.name=lang==="ar"?"مطلوب":"Required";
    const dupId=usersList.find(u=>String(u.id)===String(form.id).trim()&&String(u.id)!==String(editId));
    if(dupId)e.id=lang==="ar"?"رقم المستخدم موجود":"User ID exists";
    if(!editId&&(!form.pin||form.pin.length<4))e.pin=lang==="ar"?"PIN مطلوب (4-5 أرقام)":"PIN required (4-5 digits)";
    setErrs(e);
    return !Object.keys(e).length;
  };

  const save=()=>{
    if(!validate())return;
    const existing=editId?usersList.find(u=>String(u.id)===String(editId)):null;
    const rec={...form,id:String(form.id).trim()||String(editId),name:form.name.trim(),pin:form.pin||existing?.pin||"",permissions:Object.keys(form.permissions).length?form.permissions:DEFAULT_PERMS(form.role)};
    if(editId){persist(usersList.map(u=>String(u.id)===String(editId)?rec:u));}
    else{persist([...usersList,rec]);}
    reset();
  };

  const reset=()=>{setForm(blank);setEditId(null);setShowForm(false);setErrs({});};
  const doEdit=u=>{setForm({...u,id:u.id,pin:""});setEditId(u.id);setShowForm(true);};
  const doDelete=id=>{persist(usersList.filter(u=>u.id!==id));setDelId(null);};
  const togglePerm=(m,perm)=>{setForm(f=>{const cur=f.permissions[m]||{};return{...f,permissions:{...f.permissions,[m]:{...cur,[perm]:!cur[perm]}}};});};
  const PERM_KEYS=["view","add","edit","delete","import","export"];

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <h2 style={{fontSize:15,fontWeight:700,color:C.accent,margin:0}}>{t.users} <span style={{fontSize:12,color:C.muted}}>({usersList.length})</span></h2>
        {hasPerm(mod,"add")&&<button className="btn btn-primary" style={{fontSize:12}} onClick={()=>{setForm(blank);setShowForm(true);}}>+ {t.add}</button>}
      </div>
      <div className="table-wrap"><table>
        <thead><tr><th>{lang==="ar"?"رقم المستخدم":"User ID"}</th><th>{lang==="ar"?"الاسم":"Name"}</th><th>{lang==="ar"?"الدور":"Role"}</th><th>{lang==="ar"?"آخر دخول":"Last Login"}</th><th>{t.actions}</th></tr></thead>
        <tbody>
          {usersList.map(u=>(
            <tr key={u.id}>
              <td style={{fontWeight:600,fontFamily:"monospace"}}>{u.id}</td>
              <td style={{fontWeight:600}}>{u.name||u.username||"—"}{String(u.id)===String(currentUser?.id)&&<span style={{marginLeft:6,fontSize:10,color:C.accent}}>{lang==="ar"?"(أنت)":"(You)"}</span>}</td>
              <td><span style={{background:u.role==="admin"?"#7c3aed22":u.role==="manager"?"#3b82f622":"#22c55e22",color:u.role==="admin"?"#a78bfa":u.role==="manager"?"#60a5fa":"#22c55e",padding:"2px 8px",borderRadius:6,fontSize:11,fontWeight:700}}>{u.role}</span></td>
              <td style={{color:C.muted,fontSize:11}}>{u.lastLogin?new Date(u.lastLogin).toLocaleDateString():"-"}</td>
              <td>
                {hasPerm(mod,"edit")&&<button className="btn btn-secondary" style={{fontSize:11,padding:"3px 8px",marginRight:4}} onClick={()=>doEdit(u)}>{t.edit}</button>}
                {hasPerm(mod,"delete")&&u.id!==currentUser?.id&&<button className="btn btn-danger" style={{fontSize:11,padding:"3px 8px"}} onClick={()=>setDelId(u.id)}>{t.delete}</button>}
              </td>
            </tr>
          ))}
        </tbody>
      </table></div>

      {showForm&&<div className="overlay" onClick={e=>{if(e.target===e.currentTarget)reset();}}><div className="modal modal-lg" onClick={e=>e.stopPropagation()}>
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:14}}>{editId?t.edit:t.add} — {t.users}</h2>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10,marginBottom:14}}>
          <div><label className="lbl">{lang==="ar"?"رقم المستخدم":"User ID"}</label><input value={form.id} disabled={!!editId} onChange={e=>setForm({...form,id:e.target.value.replace(/\D/g,"")})} placeholder="1002"/>{errs.id&&<div className="err">{errs.id}</div>}</div>
          <div><label className="lbl">{lang==="ar"?"الاسم":"Name"}</label><input value={form.name||""} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">PIN{editId&&<span style={{fontSize:10,color:C.muted}}> ({lang==="ar"?"اتركه فارغ":"leave blank"})</span>}</label><input type="password" maxLength={5} value={form.pin||""} onChange={e=>setForm({...form,pin:e.target.value.replace(/\D/g,"")})}/>{errs.pin&&<div className="err">{errs.pin}</div>}</div>
          <div><label className="lbl">{lang==="ar"?"الدور":"Role"}</label>
            <select value={form.role} onChange={e=>setForm({...form,role:e.target.value,permissions:DEFAULT_PERMS(e.target.value)})}>
              <option value="admin">Admin</option>
              <option value="manager">Manager</option>
              <option value="viewer">Viewer</option>
            </select>
          </div>
        </div>
        <div className="divider"/>
        <div style={{fontWeight:700,fontSize:12,marginBottom:10,color:C.accent}}>{lang==="ar"?"الصلاحيات":"Permissions"}</div>
        <div style={{overflowX:"auto"}}>
          <table style={{minWidth:500,fontSize:11}}>
            <thead><tr>
              <th style={{textAlign:"left",padding:"4px 8px"}}>{lang==="ar"?"الوحدة":"Module"}</th>
              {PERM_KEYS.map(p=><th key={p} style={{textAlign:"center",padding:"4px 8px",textTransform:"capitalize"}}>{p}</th>)}
            </tr></thead>
            <tbody>
              {ALL_PERMS.map(m=>(
                <tr key={m}>
                  <td style={{padding:"4px 8px",fontWeight:600}}>{t[m]||m}</td>
                  {PERM_KEYS.map(p=>(
                    <td key={p} style={{textAlign:"center",padding:"4px 8px"}}>
                      <input type="checkbox" checked={!!(form.permissions[m]||{})[p]} onChange={()=>togglePerm(m,p)}/>
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{display:"flex",gap:8,marginTop:16}}><button className="btn btn-primary" style={{flex:1}} onClick={save}>{t.save}</button><button className="btn btn-secondary" style={{flex:1}} onClick={reset}>{t.cancel}</button></div>
      </div></div>}

      {delId&&<DelModal t={t} onOk={()=>doDelete(delId)} onCancel={()=>setDelId(null)}/>}
    </div>
  );
}

