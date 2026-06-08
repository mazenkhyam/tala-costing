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
    // POS / AGG pricing
    posSellPrice:"سعر POS (المحل)", aggSellPrice:"سعر AGG (التوصيل)",
    posMargin:"هامش POS", aggMargin:"هامش AGG",
    posCost:"متوسط تكلفة POS", aggCost:"متوسط تكلفة AGG",
    posAvgMargin:"متوسط هامش POS", aggAvgMargin:"متوسط هامش AGG",
    dashPOS:"لوحة POS", dashAGG:"لوحة AGG",
    pricingChannel:"قناة البيع",
    showMore:"اظهر المزيد", showLess:"اظهر أقل",
    pageSize:"عدد العناصر",
    // main categories
    mainCategory:"الفئة الرئيسية", addMainCategory:"إضافة فئة رئيسية",
    mainCatRaw:"المواد الخام", mainCatPrep:"المواد شبه المصنعة", mainCatProducts:"المنتجات",
    selectCategory:"اختر الفئة",
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
    // POS / AGG pricing
    posSellPrice:"POS Price (In-store)", aggSellPrice:"AGG Price (Delivery)",
    posMargin:"POS Margin", aggMargin:"AGG Margin",
    posCost:"Avg POS Cost", aggCost:"Avg AGG Cost",
    posAvgMargin:"Avg POS Margin", aggAvgMargin:"Avg AGG Margin",
    dashPOS:"POS Dashboard", dashAGG:"AGG Dashboard",
    pricingChannel:"Channel",
    showMore:"Show More", showLess:"Show Less",
    pageSize:"Page Size",
    mainCategory:"Main Category", addMainCategory:"Add Category",
    mainCatRaw:"Raw Materials", mainCatPrep:"Prep Items", mainCatProducts:"Products",
    selectCategory:"Select Category",
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
const defaultClasses = { raw:["Food Item","Package Item","Cleaning Item"], prep:["Sauce","Dough","Mix","Marinade"], products:["Main Dish","Beverage","Dessert","Appetizer"] };

const genCode = (prefix, existing, allLists=[]) => {
  // Collect ALL codes from existing + allLists to avoid cross-module duplicates
  const allCodes = [
    ...existing.map(i=>i.code),
    ...allLists.flatMap(l=>l.map(i=>i.code))
  ].filter(Boolean);
  const nums = allCodes
    .filter(c=>c && c.startsWith(prefix+"-"))
    .map(c=>parseInt(c.replace(prefix+"-",""))||0);
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
    const isPiece=prep.unit==="piece";
    prep.ingredients.forEach(ing=>{
      const qty=parseFloat(ing.qty)||0;
      const waste=(parseFloat(ing.waste)||0)/100;
      const netQty=qty*(1-waste); // qty after waste (goes into yield)
      if(ing.source==="prep"){
        // prep inside prep
        const subPrep=prepList.find(p=>String(p.id)===String(ing.rawId||ing.srcId||ing.prepId));
        if(!subPrep) return;
        const sub=calcPrepCostFn(subPrep,depth+1);
        // cost: if sub is piece → qty pieces × cost/piece; else qty g ÷ 1000 × cost/kg
        totalCost+=(subPrep.unit==="piece"?qty:qty/1000)*sub.costPerUnit;
      } else {
        const raw=rawList.find(r=>String(r.id)===String(ing.rawId));
        if(!raw) return;
        // cost: if raw is piece → qty pieces × price/piece; else qty g ÷ 1000 × price/kg (or liter)
        totalCost+=(raw.unit==="piece"?qty:qty/1000)*raw.price;
      }
      // yield: if prep is by piece, qty = pcs (no g conversion); else qty is in grams
      yieldG+=isPiece?netQty:netQty;
    });
    // For piece preps: yield is in pcs (1 pcs = 1 unit); divide by 1 not 1000
    // For kg/liter preps: yield is in grams; divide by 1000 to get kg/liter
    const yieldKg=prep.yieldOverride&&parseFloat(prep.yieldOverride)>0
      ?parseFloat(prep.yieldOverride)
      :isPiece
        ?yieldG          // piece: yieldG is already piece count
        :(yieldG/1000);  // kg/liter: convert from grams
    return {totalCost, yieldKg, costPerUnit:yieldKg>0?totalCost/yieldKg:0};
  };
  const calcPrepCost = useCallback((prep)=>calcPrepCostFn(prep),[rawList,prepList]);

  const calcProductCost = useCallback((prod)=>{
    if(!prod.ingredients?.length) return {totalCost:0,margin:0,posMargin:0,aggMargin:0};
    let cost=0;
    prod.ingredients.forEach(ing=>{
      const qty2=parseFloat(ing.qty)||0;
      const waste2=(parseFloat(ing.waste)||0)/100;
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
        .fbtn{background:${DARK.surface};border:1px solid ${DARK.border};color:${DARK.muted};padding:5px 12px;font-size:12px;border-radius:20px;cursor:pointer;font-family:inherit;font-weight:600;transition:all .14s}
        .fbtn.active{background:${DARK.accent}20;border-color:${DARK.accent}66;color:${DARK.accent}}
        .fbtn:hover{border-color:#3a4060;color:${DARK.text}}
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
function DashboardTab({t,lang,C=DARK,rawList,prepList,prodList,calcPrepCost,calcProductCost}) {
  const [topN,setTopN]=useState(10);
  const [section,setSection]=useState("all");
  const [channel,setChannel]=useState("pos"); // "pos" | "agg"
  const [relModal,setRelModal]=useState(null);

  const prodCalc=prodList.map(p=>({...p,...calcProductCost(p)}));
  const prepCalc=prepList.map(p=>({...p,...calcPrepCost(p)}));

  // POS / AGG split metrics
  const withPOS=prodCalc.filter(p=>parseFloat(p.posSellPrice||p.sellingPrice)>0);
  const withAGG=prodCalc.filter(p=>parseFloat(p.aggSellPrice||p.sellingPrice)>0);
  const avgPOSMargin=withPOS.length?withPOS.reduce((a,p)=>a+p.posMargin,0)/withPOS.length:0;
  const avgAGGMargin=withAGG.length?withAGG.reduce((a,p)=>a+p.aggMargin,0)/withAGG.length:0;
  const avgPOSCost=withPOS.length?withPOS.reduce((a,p)=>a+p.totalCost,0)/withPOS.length:0;
  const avgAGGCost=withAGG.length?withAGG.reduce((a,p)=>a+p.totalCost,0)/withAGG.length:0;

  // Legacy combined
  const withSP=prodCalc.filter(p=>parseFloat(p.sellingPrice||p.posSellPrice)>0);
  const avgMargin=channel==="pos"?avgPOSMargin:avgAGGMargin;
  const avgCost=channel==="pos"?avgPOSCost:avgAGGCost;
  const maxCost=prodCalc.length?Math.max(...prodCalc.map(p=>p.totalCost)):0;
  const minCost=prodCalc.filter(p=>p.totalCost>0).length?Math.min(...prodCalc.filter(p=>p.totalCost>0).map(p=>p.totalCost)):0;
  const withStd=prodCalc.filter(p=>parseFloat(p.stdCost)>0);
  const overBudget=withStd.filter(p=>p.totalCost>parseFloat(p.stdCost));

  const getMargin=(p)=>channel==="pos"?p.posMargin:p.aggMargin;
  const highM=prodCalc.filter(p=>getMargin(p)>30&&(parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0)).length;
  const midM=prodCalc.filter(p=>getMargin(p)>=15&&getMargin(p)<=30&&(parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0)).length;
  const lowM=prodCalc.filter(p=>getMargin(p)<15&&(parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0)&&getMargin(p)>0).length;
  const noPrice=prodCalc.filter(p=>!parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)).length;

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
  const atRisk=prodCalc.filter(p=>p.totalCost>0&&getMargin(p)<20&&(parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0)).sort((a,b)=>getMargin(a)-getMargin(b));

  const Bar=({val,max,color})=>(<div style={{background:C.surface,borderRadius:3,height:6,flex:1}}><div style={{width:Math.min((val/(max||1))*100,100)+"%",height:6,borderRadius:3,background:color,transition:"width .4s"}}/></div>);
  const SecHd=({c,sub})=>(<div style={{marginBottom:14}}><div style={{fontWeight:700,fontSize:14,color:C.text}}>{c}</div>{sub&&<div style={{fontSize:11,color:C.muted,marginTop:2}}>{sub}</div>}<div style={{height:2,background:C.border,marginTop:8}}/></div>);
  const KCard=({label,value,color,sub})=>(<div className="card" style={{padding:"16px 18px"}}><div style={{fontSize:26,fontWeight:800,color:color||C.accent,lineHeight:1}}>{value}</div><div style={{fontSize:12,color:C.muted,marginTop:4}}>{label}</div>{sub&&<div style={{fontSize:10,color:C.muted,marginTop:2}}>{sub}</div>}</div>);
  const noMsg=<div style={{color:C.muted,fontSize:13,textAlign:"center",padding:"24px 0"}}>{t.noData}</div>;
  const secs=[{id:"all",label:lang==="ar"?"الكل":"All"},{id:"products",label:t.secProducts},{id:"prep",label:t.secPrep},{id:"raw",label:t.secRaw},{id:"variance",label:t.varianceReport}];
  const show=s=>section==="all"||section===s;

  return (
    <div style={{maxWidth:1100,margin:"0 auto"}}>
      {/* Channel toggle + section filters */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          {secs.map(s=><button key={s.id} className={`fbtn${section===s.id?" active":""}`} onClick={()=>setSection(s.id)}>{s.label}</button>)}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:12,color:C.muted}}>{t.show}</span>
          {[5,10,20].map(n=><button key={n} className={`topn${topN===n?" active":""}`} onClick={()=>setTopN(n)}>{n}</button>)}
        </div>
      </div>

      {/* POS / AGG channel tabs */}
      <div style={{display:"flex",gap:0,marginBottom:16,background:C.surface,borderRadius:10,padding:4,width:"fit-content",border:`1px solid ${C.border}`}}>
        {[{id:"pos",label:t.dashPOS},{id:"agg",label:t.dashAGG}].map(ch=>(
          <button key={ch.id} onClick={()=>setChannel(ch.id)} style={{
            background:channel===ch.id?C.accent:"transparent",
            color:channel===ch.id?"#080b14":C.muted,
            border:"none",borderRadius:7,padding:"7px 22px",fontFamily:"inherit",
            fontSize:13,fontWeight:700,cursor:"pointer",transition:"all .16s"
          }}>{ch.label}</button>
        ))}
      </div>

      {/* KPI dual cards: POS + AGG side by side */}
      {show("products")&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12,marginBottom:20}}>
        <KCard label={lang==="ar"?"إجمالي المنتجات":"Total Products"} value={prodList.length} color={C.accent}/>
        <KCard label={lang==="ar"?"إجمالي الخام":"Total Raw"} value={rawList.length} color="#a78bfa"/>
        <KCard label={lang==="ar"?"إجمالي Prep":"Total Prep"} value={prepList.length} color={C.blue}/>
        <KCard
          label={channel==="pos"?t.posAvgMargin:t.aggAvgMargin}
          value={(channel==="pos"?avgPOSMargin:avgAGGMargin).toFixed(1)+"%"}
          color={(channel==="pos"?avgPOSMargin:avgAGGMargin)>30?C.green:(channel==="pos"?avgPOSMargin:avgAGGMargin)>15?C.yellow:C.red}
          sub={lang==="ar"?`أعلى تكلفة: ${maxCost.toFixed(2)} | أقل: ${minCost.toFixed(2)}`:`Max: ${maxCost.toFixed(2)} | Min: ${minCost.toFixed(2)}`}
        />
        <KCard label={channel==="pos"?t.posCost:t.aggCost} value={(channel==="pos"?avgPOSCost:avgAGGCost).toFixed(2)} color="#f87171"/>
        <KCard label={t.productsOverBudget} value={overBudget.length} sub={withStd.length>0?`/ ${withStd.length} ${lang==="ar"?"لهم معياري":"with target"}`:""} color={overBudget.length>0?C.red:C.green}/>
      </div>}

      {/* Side-by-side POS vs AGG summary */}
      {show("products")&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
        {[{ch:"pos",label:t.dashPOS,color:"#22c55e"},{ch:"agg",label:t.dashAGG,color:"#3b82f6"}].map(({ch,label,color})=>{
          const prods=prodCalc.filter(p=>parseFloat(ch==="pos"?(p.posSellPrice||p.sellingPrice):(p.aggSellPrice||p.sellingPrice))>0);
          const avgM=prods.length?prods.reduce((a,p)=>a+(ch==="pos"?p.posMargin:p.aggMargin),0)/prods.length:0;
          const avgC=prods.length?prods.reduce((a,p)=>a+p.totalCost,0)/prods.length:0;
          const high=prods.filter(p=>(ch==="pos"?p.posMargin:p.aggMargin)>30).length;
          const low=prods.filter(p=>(ch==="pos"?p.posMargin:p.aggMargin)<15&&(ch==="pos"?p.posMargin:p.aggMargin)>0).length;
          return (
            <div key={ch} className="card" style={{padding:"14px 16px",border:`2px solid ${channel===ch?color:C.border}`}}>
              <div style={{fontWeight:700,fontSize:13,color,marginBottom:10,display:"flex",alignItems:"center",gap:6}}>
                <span style={{width:8,height:8,borderRadius:"50%",background:color,display:"inline-block"}}/>
                {label}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                <div style={{background:C.surface,borderRadius:7,padding:"8px 10px"}}>
                  <div style={{fontSize:18,fontWeight:800,color:avgM>30?C.green:avgM>15?C.yellow:C.red}}>{avgM.toFixed(1)}%</div>
                  <div style={{fontSize:10,color:C.muted}}>{t.avgMarginLbl}</div>
                </div>
                <div style={{background:C.surface,borderRadius:7,padding:"8px 10px"}}>
                  <div style={{fontSize:18,fontWeight:800,color:"#f87171"}}>{avgC.toFixed(2)}</div>
                  <div style={{fontSize:10,color:C.muted}}>{t.avgCostLbl}</div>
                </div>
                <div style={{background:C.surface,borderRadius:7,padding:"8px 10px"}}>
                  <div style={{fontSize:16,fontWeight:700,color:C.green}}>{high}</div>
                  <div style={{fontSize:10,color:C.muted}}>{lang==="ar"?"هامش >30%":"Margin >30%"}</div>
                </div>
                <div style={{background:C.surface,borderRadius:7,padding:"8px 10px"}}>
                  <div style={{fontSize:16,fontWeight:700,color:C.red}}>{low}</div>
                  <div style={{fontSize:10,color:C.muted}}>{lang==="ar"?"هامش <15%":"Margin <15%"}</div>
                </div>
              </div>
            </div>
          );
        })}
      </div>}

      {show("products")&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.marginDistribution} sub={lang==="ar"?`توزيع المنتجات حسب هامش الربح — ${channel==="pos"?"POS":"AGG"}`:`Products by margin — ${channel==="pos"?"POS":"AGG"}`}/>
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

      {show("products")&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.secProducts}/>
        {prodCalc.length===0?noMsg:<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:16}}>
          {[
            {title:lang==="ar"?"أعلى هامش ربح":"Highest Margin",data:[...prodCalc].filter(p=>parseFloat(p.posSellPrice||p.aggSellPrice||p.sellingPrice)>0).sort((a,b)=>getMargin(b)-getMargin(a)),fn:p=>getMargin(p),fmt:v=>v.toFixed(1)+"%",col:p=>getMargin(p)>30?C.green:getMargin(p)>15?C.yellow:C.red},
            {title:lang==="ar"?"أعلى تكلفة":"Highest Cost",data:[...prodCalc].sort((a,b)=>b.totalCost-a.totalCost),fn:p=>p.totalCost,fmt:v=>v.toFixed(2),col:()=>"#f87171"},
            {title:lang==="ar"?"أقل تكلفة":"Lowest Cost",data:[...prodCalc].filter(p=>p.totalCost>0).sort((a,b)=>a.totalCost-b.totalCost),fn:p=>p.totalCost,fmt:v=>v.toFixed(2),col:()=>C.green},
          ].map((sec,si)=><div key={si}>
            <div style={{fontSize:11,color:C.muted,fontWeight:700,textTransform:"uppercase",marginBottom:10}}>{sec.title}</div>
            {sec.data.slice(0,topN).map((p,i)=>{const val=sec.fn(p),max=sec.fn(sec.data[0])||1,col=sec.col(p);return(
              <div key={p.id} style={{marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                  <span style={{fontSize:12,color:C.text,fontWeight:600,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{i+1}. {p.name}</span>
                  <span style={{fontSize:12,color:col,fontWeight:700,marginRight:6}}>{sec.fmt(val)}</span>
                </div>
                <Bar val={val} max={max} color={col}/>
              </div>
            );})}
          </div>)}
        </div>}
      </div>}

      {show("products")&&atRisk.length>0&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={lang==="ar"?"منتجات تحتاج مراجعة (هامش أقل من 20%)":"Products Needing Review (Margin < 20%)"} sub={lang==="ar"?"تحقق من التسعير أو خفض التكلفة":"Check pricing or reduce costs"}/>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>{[t.productName,t.totalCost,t.sellingPrice,t.margin,t.topCostDriver].map((h,i)=><th key={i} style={{padding:"9px 12px",textAlign:lang==="ar"?"right":"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",background:C.surface,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>{atRisk.slice(0,topN).map((p,i)=>{
              let topDrv="—",topC=0;
              p.ingredients?.forEach(ing=>{let cost=0,name="";if(ing.source==="raw"){const r=rawList.find(r=>String(r.id)===String(ing.srcId));if(r){cost=(parseFloat(ing.qty)||0)/1000*r.price;name=r.name;}}else{const pr=prepList.find(pr=>String(pr.id)===String(ing.srcId));if(pr){const {costPerUnit}=calcPrepCost(pr);cost=(parseFloat(ing.qty)||0)/1000*costPerUnit;name=pr.name;}}if(cost>topC){topC=cost;topDrv=name;}});
              return <tr key={p.id} style={{borderBottom:`1px solid ${C.border}22`}}>
                <td style={{padding:"10px 12px",fontWeight:600}}>{p.name}</td>
                <td style={{padding:"10px 12px",color:"#f87171",fontWeight:600}}>{p.totalCost.toFixed(2)}</td>
                <td style={{padding:"10px 12px",color:C.green,fontWeight:600}}>{parseFloat(p.sellingPrice||0).toFixed(2)}</td>
                <td style={{padding:"10px 12px"}}><span style={{background:p.margin<0?"#f8717122":"#f59e0b22",color:p.margin<0?"#f87171":"#f59e0b",padding:"2px 8px",borderRadius:20,fontSize:12,fontWeight:700}}>{p.margin.toFixed(1)}%</span></td>
                <td style={{padding:"10px 12px",color:C.yellow,fontSize:12,fontWeight:600}}>{topDrv} <span style={{color:C.muted,fontSize:11}}>({topC.toFixed(2)})</span></td>
              </tr>;
            })}</tbody>
          </table>
        </div>
      </div>}

      {show("variance")&&withStd.length>0&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.varianceReport} sub={lang==="ar"?"مقارنة التكلفة الفعلية بالمعيارية":"Actual vs Standard Cost"}/>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>{[t.productName,t.totalCost,t.stdCost,t.variance,t.variancePct,t.topCostDriver].map((h,i)=><th key={i} style={{padding:"9px 12px",textAlign:lang==="ar"?"right":"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",background:C.surface,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>{[...withStd].sort((a,b)=>(b.totalCost-parseFloat(b.stdCost))-(a.totalCost-parseFloat(a.stdCost))).slice(0,topN).map((p,i)=>{
              const v=p.totalCost-parseFloat(p.stdCost); const vp=(v/parseFloat(p.stdCost))*100;
              let topDrv="—",topC=0;
              p.ingredients?.forEach(ing=>{let cost=0,name="";if(ing.source==="raw"){const r=rawList.find(r=>String(r.id)===String(ing.srcId));if(r){cost=(parseFloat(ing.qty)||0)/1000*r.price;name=r.name;}}else{const pr=prepList.find(pr=>String(pr.id)===String(ing.srcId));if(pr){const {costPerUnit}=calcPrepCost(pr);cost=(parseFloat(ing.qty)||0)/1000*costPerUnit;name=pr.name;}}if(cost>topC){topC=cost;topDrv=name;}});
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

      {show("products")&&<div className="card" style={{padding:18,marginBottom:16}}>
        <SecHd c={t.topCostDriversReport} sub={lang==="ar"?"أعلى المكونات تكلفة عبر جميع المنتجات":"Top cost ingredients across all products"}/>
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
              <button onClick={()=>p.usedIn>0&&setRelModal({title:p.name,list:p.relatedProds,type:"products"})}
                style={{background:p.usedIn>0?C.blue+"22":"transparent",color:p.usedIn>0?C.blue:C.muted,padding:"2px 10px",borderRadius:20,fontSize:12,fontWeight:700,border:"none",cursor:p.usedIn>0?"pointer":"default",fontFamily:"inherit"}}>
                {p.usedIn} {t.productsCount}
              </button>
            </div>)}
          </div>
        </div>}
      </div>}

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
                  <button onClick={()=>r.inPrep>0&&setRelModal({title:r.name,list:r.relatedPreps,type:"prep"})}
                    style={{background:r.inPrep>0?"#1e3a5f":"transparent",color:r.inPrep>0?"#60a5fa":C.muted,padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700,border:"none",cursor:r.inPrep>0?"pointer":"default",fontFamily:"inherit"}}>P:{r.inPrep}</button>
                  <button onClick={()=>r.inProd>0&&setRelModal({title:r.name,list:r.relatedProds,type:"products"})}
                    style={{background:r.inProd>0?"#0a3326":"transparent",color:r.inProd>0?"#4ade80":C.muted,padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700,border:"none",cursor:r.inProd>0?"pointer":"default",fontFamily:"inherit"}}>M:{r.inProd}</button>
                </div>
              </div>
            </div>)}
          </div>
        </div>}
      </div>}

      {relModal&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&setRelModal(null)}>
        <div className="modal" style={{maxWidth:500}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
            <div><div style={{fontWeight:800,fontSize:15,color:C.accent}}>{relModal.title}</div>
              <div style={{fontSize:12,color:C.muted,marginTop:2}}>{relModal.type==="prep"?t.relatedPreps:t.relatedProducts} ({relModal.list.length})</div>
            </div>
            <button onClick={()=>setRelModal(null)} style={{background:"transparent",border:"none",color:C.muted,fontSize:20,cursor:"pointer"}}>✕</button>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {relModal.list.map((item,i)=><div key={item.id||i} style={{background:C.surface,borderRadius:9,padding:"10px 14px",border:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><code style={{background:C.bg||C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8",marginLeft:8}}>{item.code}</code><span style={{fontWeight:600,fontSize:13}}>{item.name}</span></div>
              <span className="badge badge-cls">{item.class||"—"}</span>
            </div>)}
          </div>
          <div style={{display:"flex",justifyContent:"flex-end",marginTop:14}}><button className="btn btn-secondary" onClick={()=>setRelModal(null)}>{t.cancel}</button></div>
        </div>
      </div>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// PREVIEW MODAL — shown after file parse, before commit
// ═══════════════════════════════════════════════════════════════
function PreviewModal({t,lang,C=DARK,data,onConfirm,onCancel,classes,moduleType}) {
  const [items,setItems]=useState(()=>data.items.map((it,i)=>({...it,_idx:i,_skip:false})));
  const ar=lang==="ar";
  const unitOpts=["kg","liter","piece"];
  const clsOpts=moduleType==="raw"?(classes||[]):moduleType==="prep"?(classes||[]):[];
  const addCount=items.filter(it=>!it._skip&&!it._existing).length;
  const updCount=items.filter(it=>!it._skip&&it._existing).length;
  const skipCount=items.filter(it=>it._skip).length;
  const update=(idx,field,val)=>setItems(prev=>prev.map(it=>it._idx===idx?{...it,[field]:val}:it));
  return (
    <div className="overlay" onClick={e=>e.target===e.currentTarget&&onCancel()}>
      <div className="modal" style={{maxWidth:900,width:"100%",maxHeight:"90vh",display:"flex",flexDirection:"column"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexShrink:0}}>
          <div>
            <div style={{fontWeight:800,fontSize:15,color:C.accent}}>{ar?"مراجعة البيانات قبل الاستيراد":"Review Before Import"}</div>
            <div style={{fontSize:12,color:C.muted,marginTop:3,display:"flex",gap:12}}>
              <span style={{color:C.green}}>✓ {ar?"يضاف":"Add"}: {addCount}</span>
              <span style={{color:C.yellow}}>↻ {ar?"يحدَّث":"Update"}: {updCount}</span>
              <span style={{color:C.muted}}>⊘ {ar?"يتجاهل":"Skip"}: {skipCount}</span>
            </div>
          </div>
          <button onClick={onCancel} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer"}}>✕</button>
        </div>
        <div style={{overflowY:"auto",flex:1,marginBottom:14}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead style={{position:"sticky",top:0,zIndex:10}}>
              <tr>
                {[ar?"#":"#",ar?"الاسم":"Name",ar?"الكلاس":"Class",ar?"الوحدة":"Unit",ar?"السعر":"Price",ar?"الإجراء":"Action",ar?"تجاهل":"Skip"].map((h,i)=>(
                  <th key={i} style={{padding:"8px 10px",textAlign:ar?"right":"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",background:C.surface,borderBottom:`1px solid ${C.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {items.map((it)=>(
                <tr key={it._idx} style={{opacity:it._skip?0.4:1,background:it._skip?"transparent":it._existing?C.accent+"08":C.green+"08"}}>
                  <td style={{padding:"7px 10px",color:C.muted,fontSize:11}}>{it._idx+1}</td>
                  <td style={{padding:"7px 10px"}}>
                    <input value={it.name} onChange={e=>update(it._idx,"name",e.target.value)}
                      style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"4px 8px",fontSize:12,width:"100%",outline:"none",minWidth:120}}/>
                  </td>
                  <td style={{padding:"7px 10px"}}>
                    {clsOpts.length>0
                      ?<select value={it.class||""} onChange={e=>update(it._idx,"class",e.target.value)} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"4px 8px",fontSize:12,width:"100%",outline:"none"}}>
                        <option value="">—</option>
                        {clsOpts.map(c=><option key={c} value={c}>{c}</option>)}
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
                      ?<input type="number" value={it.price||0} onChange={e=>update(it._idx,"price",e.target.value)}
                          style={{background:C.surface,border:`1px solid ${C.border}`,color:C.accent,borderRadius:5,padding:"4px 8px",fontSize:12,width:80,outline:"none",fontWeight:700}}/>
                      :<span style={{color:C.muted}}>—</span>}
                  </td>
                  <td style={{padding:"7px 10px"}}>
                    <span style={{padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700,
                      background:it._existing?C.yellow+"22":C.green+"22",
                      color:it._existing?C.yellow:C.green}}>
                      {it._existing?(ar?"تحديث":"Update"):(ar?"إضافة":"Add")}
                    </span>
                  </td>
                  <td style={{padding:"7px 10px",textAlign:"center"}}>
                    <input type="checkbox" checked={!!it._skip} onChange={e=>update(it._idx,"_skip",e.target.checked)}
                      style={{width:15,height:15,accentColor:C.danger,cursor:"pointer"}}/>
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
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// DUPLICATE DELETE MODAL — manage duplicate names
// ═══════════════════════════════════════════════════════════════
function DupDeleteModal({t,lang,C=DARK,items,onDelete,onClose}) {
  const ar=lang==="ar";
  // Group by name
  const groups={};
  items.forEach(it=>{
    const k=it.name.trim().toLowerCase();
    if(!groups[k]) groups[k]=[];
    groups[k].push(it);
  });
  const [toDelete,setToDelete]=useState({});
  const toggleDel=(id)=>setToDelete(p=>({...p,[id]:!p[id]}));
  const selectedCount=Object.values(toDelete).filter(Boolean).length;
  return (
    <div className="overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="modal" style={{maxWidth:680,width:"100%",maxHeight:"85vh",display:"flex",flexDirection:"column"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexShrink:0}}>
          <div>
            <div style={{fontWeight:800,fontSize:15,color:C.yellow}}>{ar?"إدارة الأسماء المكررة":"Manage Duplicate Names"}</div>
            <div style={{fontSize:12,color:C.muted,marginTop:3}}>{ar?"حدد السجلات التي تريد حذفها":"Select the records to delete"}</div>
          </div>
          <button onClick={onClose} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer"}}>✕</button>
        </div>
        <div style={{overflowY:"auto",flex:1,marginBottom:14}}>
          {Object.entries(groups).map(([key,grp])=>(
            <div key={key} style={{marginBottom:14,background:C.surface,borderRadius:10,padding:12,border:`1px solid ${C.yellow}44`}}>
              <div style={{fontSize:12,fontWeight:700,color:C.yellow,marginBottom:8}}>
                ⚠ {ar?"اسم مكرر":"Duplicate"}: <span style={{color:C.text}}>{grp[0].name}</span>
                <span style={{color:C.muted,fontWeight:400,marginRight:6,marginLeft:6}}>({grp.length} {ar?"نسخ":"copies"})</span>
              </div>
              {grp.map((it,i)=>(
                <div key={it.id} style={{display:"flex",alignItems:"center",gap:10,padding:"7px 10px",borderRadius:7,background:toDelete[it.id]?C.danger+"18":C.card,border:`1px solid ${toDelete[it.id]?C.danger+"66":C.border}`,marginBottom:5}}>
                  <input type="checkbox" checked={!!toDelete[it.id]} onChange={()=>toggleDel(it.id)} style={{width:14,height:14,accentColor:C.danger,cursor:"pointer"}}/>
                  <code style={{background:C.bg||C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{it.code}</code>
                  <span style={{fontWeight:600,fontSize:12,flex:1}}>{it.name}</span>
                  <span className="badge badge-cls">{it.class||"—"}</span>
                  <span style={{fontSize:11,color:C.muted}}>{it.unit}</span>
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
            🗑 {ar?`حذف ${selectedCount} سجل محدد`:`Delete ${selectedCount} selected`}
          </button>
          <button className="btn btn-secondary" style={{flex:1}} onClick={onClose}>{t.cancel}</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// RAW MATERIALS
// ═══════════════════════════════════════════════════════════════
// ─── USAGE MODAL ─────────────────────────────────────────────
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
          <div><code style={{background:C.bg||C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8",marginLeft:8}}>{item.code}</code><span style={{fontWeight:600,fontSize:13}}>{item.name}</span></div>
          <span className="badge badge-cls">{item.class||"—"}</span>
        </div>)}
      </div>
      <div style={{display:"flex",justifyContent:"flex-end",marginTop:14}}><button className="btn btn-secondary" onClick={onClose}>{t.cancel}</button></div>
    </div>
  </div>;
}

// ─── IMPORT MODAL ─────────────────────────────────────────────
function ImportModal({t,lang,C=DARK,type,onClose,onFileSelect,onDownloadTemplate}) {
  const fileRef=useRef();
  return (
    <div className="overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="modal" style={{maxWidth:420}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div>
            <div style={{fontWeight:800,fontSize:15,color:C.accent}}>{lang==="ar"?"استيراد البيانات":"Import Data"}</div>
            <div style={{fontSize:12,color:C.muted,marginTop:2}}>{type}</div>
          </div>
          <button onClick={onClose} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer",lineHeight:1}}>✕</button>
        </div>
        <div onClick={onDownloadTemplate} style={{background:C.surface,border:`2px solid ${C.accent}44`,borderRadius:12,padding:"18px 20px",marginBottom:12,cursor:"pointer",transition:"all .18s",display:"flex",alignItems:"center",gap:14}}
          onMouseOver={e=>{e.currentTarget.style.borderColor=C.accent;e.currentTarget.style.background=C.accent+"11";}}
          onMouseOut={e=>{e.currentTarget.style.borderColor=C.accent+"44";e.currentTarget.style.background=C.surface;}}>
          <div style={{fontSize:28,flexShrink:0}}>📥</div>
          <div>
            <div style={{fontWeight:700,color:C.accent,fontSize:13,marginBottom:3}}>{lang==="ar"?"تحميل النموذج الفارغ":"Download Empty Template"}</div>
            <div style={{fontSize:11,color:C.muted}}>{lang==="ar"?"احصل على النموذج الصحيح وعبّأه بالبيانات":"Get the correct format and fill it with your data"}</div>
          </div>
        </div>
        <div onClick={()=>fileRef.current.click()} style={{background:C.surface,border:`2px solid ${C.blue}44`,borderRadius:12,padding:"18px 20px",cursor:"pointer",transition:"all .18s",display:"flex",alignItems:"center",gap:14}}
          onMouseOver={e=>{e.currentTarget.style.borderColor=C.blue;e.currentTarget.style.background=C.blue+"11";}}
          onMouseOut={e=>{e.currentTarget.style.borderColor=C.blue+"44";e.currentTarget.style.background=C.surface;}}>
          <div style={{fontSize:28,flexShrink:0}}>📤</div>
          <div>
            <div style={{fontWeight:700,color:C.blue,fontSize:13,marginBottom:3}}>{lang==="ar"?"رفع ملف Excel":"Upload Excel File"}</div>
            <div style={{fontSize:11,color:C.muted}}>{lang==="ar"?"ارفع ملف Excel مملوء بالبيانات":"Upload a filled Excel file"}</div>
          </div>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{onFileSelect(e);onClose();}}/>
        <div style={{marginTop:14,padding:"10px 14px",background:"#060810",borderRadius:8,fontSize:11,color:C.muted}}>💡 {lang==="ar"?"النظام يتعرف تلقائياً على أي تنسيق — لا يهم أسماء الأعمدة":"Auto-detects any column format"}</div>
        <button className="btn btn-secondary" onClick={onClose} style={{width:"100%",marginTop:12,textAlign:"center"}}>{t.cancel}</button>
      </div>
    </div>
  );
}

// ─── INGREDIENT ROW (PREP) ────────────────────────────────────
function IngRow({ing,rawList,prepList,lang,t,C=DARK,onUpdate,onRemove}) {
  const [q,setQ]=useState(""); const [open,setOpen]=useState(false);
  const srcList=ing.source==="prep"?(prepList||[]):(rawList||[]);
  const filtered=srcList.filter(r=>r.name.toLowerCase().includes(q.toLowerCase())||r.code?.toLowerCase().includes(q.toLowerCase()));
  const selected=srcList.find(r=>String(r.id)===String(ing.rawId));
  return (
    <div style={{display:"grid",gap:7,gridTemplateColumns:"1fr 2fr 1fr 1fr auto",alignItems:"end",marginBottom:7,background:C.surface,padding:9,borderRadius:8,border:`1px solid ${C.border}`}}>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.source}</label>
        <select value={ing.source||"raw"} onChange={e=>{onUpdate("source",e.target.value);onUpdate("rawId","");}} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}>
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
          <div style={{padding:6}}><input autoFocus placeholder={lang==="ar"?"بحث...":"Search..."} value={q} onChange={e=>setQ(e.target.value)} onClick={e=>e.stopPropagation()} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"6px 10px",fontSize:11,outline:"none",width:"100%"}}/></div>
          <div style={{maxHeight:200,overflowY:"auto"}}>
            <div onClick={()=>{onUpdate("rawId","");setOpen(false);setQ("");}} style={{padding:"7px 10px",cursor:"pointer",fontSize:12,color:C.muted}}>—</div>
            {filtered.map(r=><div key={r.id} onClick={()=>{onUpdate("rawId",r.id);setOpen(false);setQ("");}} style={{padding:"7px 10px",cursor:"pointer",fontSize:12,color:String(r.id)===String(ing.rawId)?C.accent:C.text,background:String(r.id)===String(ing.rawId)?C.accent+"15":"transparent"}}>{r.name} <span style={{color:C.muted,fontSize:10}}>({r.code})</span></div>)}
            {filtered.length===0&&<div style={{padding:"7px 10px",color:C.muted,fontSize:11}}>{t.noData}</div>}
          </div>
        </div>}
      </div>
      <div><label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.qty} g</label><input type="number" min="0" step="1" value={ing.qty} onChange={e=>onUpdate("qty",e.target.value)} placeholder="0" style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}/></div>
      <div><label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.waste} %</label><input type="number" min="0" max="100" step="0.5" value={ing.waste} onChange={e=>onUpdate("waste",e.target.value)} placeholder="0" style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}/></div>
      <button onClick={onRemove} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"4px 10px",cursor:"pointer",marginTop:22,fontFamily:"inherit",fontWeight:600,fontSize:12}}>✕</button>
    </div>
  );
}

// ─── INGREDIENT ROW (PRODUCT) ─────────────────────────────────
function IngRowProd({ing,rawList,prepList,lang,t,C=DARK,onUpdate,onRemove}) {
  const [q,setQ]=useState(""); const [open,setOpen]=useState(false);
  const srcList=ing.source==="raw"?rawList:prepList;
  const filtered=srcList.filter(r=>r.name.toLowerCase().includes(q.toLowerCase())||r.code?.toLowerCase().includes(q.toLowerCase()));
  const selected=srcList.find(r=>String(r.id)===String(ing.srcId));
  return (
    <div style={{display:"grid",gap:7,gridTemplateColumns:"1fr 2fr 1fr 1fr auto",alignItems:"end",marginBottom:7,background:C.surface,padding:9,borderRadius:8,border:`1px solid ${C.border}`}}>
      <div>
        <label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.source}</label>
        <select value={ing.source||"raw"} onChange={e=>{onUpdate("source",e.target.value);onUpdate("srcId","");}} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}>
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
          <div style={{padding:6}}><input autoFocus placeholder={lang==="ar"?"بحث...":"Search..."} value={q} onChange={e=>setQ(e.target.value)} onClick={e=>e.stopPropagation()} style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:5,padding:"6px 10px",fontSize:11,outline:"none",width:"100%"}}/></div>
          <div style={{maxHeight:200,overflowY:"auto"}}>
            <div onClick={()=>{onUpdate("srcId","");setOpen(false);setQ("");}} style={{padding:"7px 10px",cursor:"pointer",fontSize:12,color:C.muted}}>—</div>
            {filtered.map(r=><div key={r.id} onClick={()=>{onUpdate("srcId",r.id);setOpen(false);setQ("");}} style={{padding:"7px 10px",cursor:"pointer",fontSize:12,color:String(r.id)===String(ing.srcId)?C.accent:C.text,background:String(r.id)===String(ing.srcId)?C.accent+"15":"transparent"}}>{r.name} <span style={{color:C.muted,fontSize:10}}>({r.code})</span></div>)}
            {filtered.length===0&&<div style={{padding:"7px 10px",color:C.muted,fontSize:11}}>{t.noData}</div>}
          </div>
        </div>}
      </div>
      <div><label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.qty} g</label><input type="number" min="0" step="1" value={ing.qty} onChange={e=>onUpdate("qty",e.target.value)} placeholder="0" style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}/></div>
      <div><label style={{fontSize:11,color:C.muted,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase"}}>{t.waste} %</label><input type="number" min="0" max="100" step="0.5" value={ing.waste} onChange={e=>onUpdate("waste",e.target.value)} placeholder="0" style={{background:C.surface,border:`1px solid ${C.border}`,color:C.text,borderRadius:7,padding:"8px 10px",fontSize:12,width:"100%",outline:"none"}}/></div>
      <button onClick={onRemove} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"4px 10px",cursor:"pointer",marginTop:22,fontFamily:"inherit",fontWeight:600,fontSize:12}}>✕</button>
    </div>
  );
}


function RawTab({t,lang,C=DARK,rawList,setRawList,classes,prepList=[],prodList=[],showToast,hasPerm,mod}) {
  const [searchRaw,setSearchRaw]=useState("");
  const [search,setSearch]=useState("");
  const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false);
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",unit:"kg",price:"",class:""});
  const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null);
  const [usageModal,setUsageModal]=useState(null);
  const [showImport,setShowImport]=useState(false);
  const [previewData,setPreviewData]=useState(null);
  const [bulkDelModal,setBulkDelModal]=useState(null);
  const [pageSize,setPageSize]=useState(20);
  const [showCount,setShowCount]=useState(20);
  const fileRef=useRef();
  const cls=classes.raw||[];

  // Debounce search for speed
  useEffect(()=>{const timer=setTimeout(()=>setSearch(searchRaw),200);return()=>clearTimeout(timer);},[searchRaw]);
  // Reset showCount when filter/search changes
  useEffect(()=>setShowCount(pageSize),[search,fcls,pageSize]);

  const ok=()=>{ const e={}; if(!form.name.trim())e.name=t.required; if(!form.price||parseFloat(form.price)<=0)e.price=t.positiveNum; setErrs(e); return !Object.keys(e).length; };
  const reset=()=>{ setForm({name:"",unit:"kg",price:"",class:""}); setErrs({}); setShowForm(false); setEditId(null); };
  const save=()=>{
    if(!ok()) return;
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    if(editId!==null) setRawList(p=>p.map(m=>m.id===editId?{...m,name:form.name.trim(),unit:form.unit,price:parseFloat(form.price),class:form.class,lastUpdated:now}:m));
    else { const code=genCode("Raw",rawList); setRawList(p=>[...p,{id:Date.now(),code,name:form.name.trim(),unit:form.unit,price:parseFloat(form.price),class:form.class,lastUpdated:now}]); }
    reset(); showToast(t.savedOk);
  };
  const doEdit=m=>{ setForm({name:m.name,unit:m.unit,price:String(m.price),class:m.class||""}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setRawList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };

  const findDuplicates=()=>{
    const nameMap={};
    rawList.forEach(m=>{
      const k=m.name.trim().toLowerCase();
      if(!nameMap[k]) nameMap[k]=[];
      nameMap[k].push(m);
    });
    return Object.values(nameMap).filter(g=>g.length>1).flat();
  };
  const dups=findDuplicates();

  const doDownloadTemplate=()=>{
    const sample=[
      {Code:"Raw-00001",Name:lang==="ar"?"دجاج خام":"Chicken Raw",Class:"Food Item",Unit:"kg",Price:20},
      {Code:"Raw-00002",Name:lang==="ar"?"طماطم":"Tomato",Class:"Food Item",Unit:"kg",Price:5},
      {Code:"Raw-00003",Name:lang==="ar"?"زيت نباتي":"Vegetable Oil",Class:"Food Item",Unit:"liter",Price:8},
      {Code:"",Name:"",Class:"",Unit:"",Price:""},
    ];
    const ws=XLSX.utils.json_to_sheet(sample);
    ws["!cols"]=[{wch:14},{wch:30},{wch:14},{wch:10},{wch:12}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Raw Materials");
    XLSX.writeFile(wb,"TALA_Raw_Materials_Template.xlsx");
  };

  const doExport=()=>{
    const data=rawList.map(m=>({Code:m.code,Name:m.name,Class:m.class||"",Unit:m.unit,Price:m.price}));
    if(!data.length){doDownloadTemplate();return;}
    const ws=XLSX.utils.json_to_sheet(data);
    ws["!cols"]=[{wch:14},{wch:30},{wch:14},{wch:10},{wch:12}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Raw Materials");
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
          const rpRaw=fv(row,["Price","السعر","price","Cost","التكلفة","تكلفة","سعر","Cost Per KG","التكلفة بالكيلو","New Cost"]);
          const rp=parseFloat(rpRaw)||0;
          const ruRaw=fv(row,["Unit","الوحدة","unit","وحدة","UOM"]);
          const ru=ruRaw.toLowerCase();
          const rcls=fv(row,["Class","الكلاس","class","كلاس","Category","فئة","Type","النوع"])||"Food Item";
          if(!rn) return;
          const unit=["liter","litre","l","لتر"].includes(ru)?"liter":["piece","pieces","pcs","حبة","قطعة"].includes(ru)?"piece":"kg";
          const existing=rawList.find(m=>m.name.toLowerCase()===rn.toLowerCase()) ||
                         (rc?rawList.find(m=>m.code===rc):null);
          items.push({name:rn,code:rc||"",unit,price:rp,class:rcls,_existing:existing,_action:existing?"update":"add"});
        });
        setPreviewData({items, module:"raw"});
        setShowImport(false);
      }catch(err){showToast(lang==="ar"?"خطأ في الملف":"File error","error");}
    };
    r.readAsBinaryString(file); e.target.value="";
  };

  const confirmImport=(editedItems)=>{
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    let updated=0,added=0;
    setRawList(prev=>{
      const u=[...prev];
      editedItems.forEach(item=>{
        if(item._skip) return;
        const i=item._existing?u.findIndex(m=>m.id===item._existing.id):-1;
        if(i!==-1){
          u[i]={...u[i],name:item.name.trim(),price:parseFloat(item.price)||0,unit:item.unit,class:item.class||u[i].class,lastUpdated:now};
          updated++;
        } else {
          const newCode=item.code&&!["undefined","nan",""].includes(item.code)?item.code:genCode("Raw",u);
          u.push({id:Date.now()+Math.random(),code:newCode,name:item.name.trim(),unit:item.unit,price:parseFloat(item.price)||0,class:item.class||"Food Item",lastUpdated:now});
          added++;
        }
      });
      return u;
    });
    setPreviewData(null);
    showToast(lang==="ar"?`✅ أضيف ${added} | حُدِّث ${updated}`:`✅ Added ${added} | Updated ${updated}`,(added+updated)>0?"success":"warning");
  };

  const [selected,setSelected]=useState(new Set());
  const toggleSel=(id)=>setSelected(p=>{const n=new Set(p);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>setSelected(p=>p.size===displayed.length?new Set():new Set(displayed.map(m=>m.id)));
  const bulkDelete=()=>{setRawList(p=>p.filter(m=>!selected.has(m.id)));setSelected(new Set());showToast(lang==="ar"?`تم حذف ${selected.size} صنف`:`Deleted ${selected.size} items`,"error");};

  const filtered=rawList.filter(m=>
    (m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase()))
    &&(fcls==="all"||m.class===fcls)
  );
  const displayed=filtered.slice(0,showCount);
  const hasMore=filtered.length>showCount;

  return (
    <div>
      {/* Duplicate warning banner */}
      {dups.length>0&&<div style={{background:"#3d2205",border:"1px solid #ca8a04",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:16}}>⚠️</span>
          <span style={{fontSize:12,color:"#fcd34d",fontWeight:700}}>
            {lang==="ar"?`يوجد ${dups.length} مادة خام بأسماء مكررة`:`${dups.length} duplicate raw material names found`}
          </span>
        </div>
        <button onClick={()=>setBulkDelModal({items:dups,module:"raw"})}
          style={{background:"#ca8a04",color:"#060810",border:"none",borderRadius:6,padding:"5px 12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
          {lang==="ar"?"عرض وإدارة التكرارات":"View & Manage Duplicates"}
        </button>
      </div>}

      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:240}} placeholder={t.search} value={searchRaw} onChange={e=>{setSearchRaw(e.target.value);setSelected(new Set());}}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
          {/* Page size selector */}
          <div style={{display:"flex",alignItems:"center",gap:5,background:C.surface,borderRadius:7,padding:"4px 8px",border:`1px solid ${C.border}`}}>
            <span style={{fontSize:11,color:C.muted}}>{t.pageSize}:</span>
            {[5,10,20].map(n=><button key={n} onClick={()=>{setPageSize(n);setShowCount(n);}} style={{background:pageSize===n?C.accent:"transparent",color:pageSize===n?"#080b14":C.muted,border:"none",borderRadius:5,padding:"3px 8px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{n}</button>)}
          </div>
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={()=>setShowImport(true)}>📥 {t.importXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>

      {/* Bulk action bar */}
      {selected.size>0&&hasPerm(mod,"delete")&&<div style={{background:"#1a0f0f",border:"1px solid #dc262666",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <span style={{fontSize:13,color:"#fca5a5",fontWeight:700}}>
          {lang==="ar"?`تم تحديد ${selected.size} صنف`:`${selected.size} item(s) selected`}
        </span>
        <div style={{display:"flex",gap:7}}>
          <button onClick={()=>setSelected(new Set())} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:6,padding:"5px 12px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>
            {lang==="ar"?"إلغاء التحديد":"Deselect"}
          </button>
          <button onClick={bulkDelete} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"5px 14px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>
            🗑 {lang==="ar"?`حذف ${selected.size} محدد`:`Delete ${selected.size}`}
          </button>
        </div>
      </div>}

      <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <button className={`fbtn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all} ({rawList.length})</button>
        {cls.map(c=><button key={c} className={`fbtn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c} ({rawList.filter(m=>m.class===c).length})</button>)}
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>
            {hasPerm(mod,"delete")&&<th style={{width:36,padding:"10px 8px"}}>
              <input type="checkbox" checked={displayed.length>0&&selected.size===displayed.length} onChange={toggleAll} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/>
            </th>}
            {["#",t.code,t.name,t.class,t.unit,t.price,t.usedInPrep,t.usedInProducts,hasPerm(mod,"edit")||hasPerm(mod,"delete")?t.actions:""].filter(Boolean).map((h,i)=><th key={i}>{h}</th>)}
          </tr></thead>
          <tbody>
            {displayed.length===0?<tr><td colSpan={hasPerm(mod,"delete")?10:9} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :displayed.map((m,i)=>{
              const nPrep=prepList.filter(p=>p.ingredients?.some(i=>String(i.rawId)===String(m.id))).length;
              const nProd=prodList.filter(p=>p.ingredients?.some(i=>i.source==="raw"&&String(i.srcId)===String(m.id))).length;
              const isDup=dups.some(d=>d.id===m.id);
              const isSel=selected.has(m.id);
              return <tr key={m.id} style={{background:isSel?C.accent+"12":isDup?"#3d220511":"",cursor:"default"}}>
                {hasPerm(mod,"delete")&&<td style={{padding:"10px 8px",textAlign:"center"}}>
                  <input type="checkbox" checked={isSel} onChange={()=>toggleSel(m.id)} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/>
                </td>}
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>
                  {isDup&&<span title={lang==="ar"?"اسم مكرر":"Duplicate name"} style={{color:"#f59e0b",marginLeft:4,marginRight:4,fontSize:12}}>⚠</span>}
                  {m.name}
                </td>
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
      {/* Show more / pagination */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginTop:10}}>
        <span style={{fontSize:12,color:C.muted}}>{lang==="ar"?`عرض ${displayed.length} من ${filtered.length}`:`Showing ${displayed.length} of ${filtered.length}`}</span>
        {hasMore&&<button onClick={()=>setShowCount(c=>c+pageSize)} style={{background:C.surface,color:C.accent,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{t.showMore} (+{Math.min(pageSize,filtered.length-showCount)})</button>}
        {showCount>pageSize&&!hasMore&&<button onClick={()=>setShowCount(pageSize)} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{t.showLess}</button>}
      </div>

      {showImport&&<ImportModal t={t} lang={lang} C={C} type={t.rawMat} onClose={()=>setShowImport(false)} onDownloadTemplate={doDownloadTemplate} onFileSelect={doImport}/>}
      {usageModal&&<UsageModal t={t} C={C} usageModal={usageModal} onClose={()=>setUsageModal(null)}/>}
      {previewData&&previewData.module==="raw"&&<PreviewModal t={t} lang={lang} C={C} data={previewData} onConfirm={confirmImport} onCancel={()=>setPreviewData(null)} classes={cls} moduleType="raw"/>}
      {bulkDelModal&&bulkDelModal.module==="raw"&&<DupDeleteModal t={t} lang={lang} C={C} items={bulkDelModal.items} onDelete={ids=>{setRawList(p=>p.filter(m=>!ids.includes(m.id)));setBulkDelModal(null);showToast(lang==="ar"?`تم حذف ${ids.length} مادة مكررة`:`Deleted ${ids.length} duplicates`,"error");}} onClose={()=>setBulkDelModal(null)}/>}
      {showForm&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&reset()}><div className="modal">
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



function PrepTab({t,lang,C=DARK,prepList,setPrepList,rawList,prodList=[],classes,calcPrepCost,showToast,hasPerm,mod}) {
  const [search,setSearch]=useState(""); const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",unit:"kg",class:"",yieldOverride:"",ingredients:[]}); const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null); const [viewItem,setViewItem]=useState(null);
  const [usageModal,setUsageModal]=useState(null); const [showImport,setShowImport]=useState(false);
  const [previewData,setPreviewData]=useState(null);
  const [bulkDelModal,setBulkDelModal]=useState(null);
  const fileRef=useRef(); const cls=classes.prep||[];

  // Find duplicates
  const findDuplicates=()=>{
    const nameMap={};
    prepList.forEach(m=>{const k=m.name.trim().toLowerCase();if(!nameMap[k])nameMap[k]=[];nameMap[k].push(m);});
    return Object.values(nameMap).filter(g=>g.length>1).flat();
  };
  const dups=findDuplicates();

  const blank=()=>({id:Date.now()+Math.random(),source:"raw",rawId:"",qty:"",waste:"0"});
  const reset=()=>{ setForm({name:"",unit:"kg",class:"",yieldOverride:"",ingredients:[]}); setErrs({}); setShowForm(false); setEditId(null); };
  const ok=()=>{ const e={}; if(!form.name.trim())e.name=t.required; setErrs(e); return !Object.keys(e).length; };
  const save=()=>{
    if(!ok()) return;
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    const ings=form.ingredients.filter(i=>i.rawId&&parseFloat(i.qty)>0);
    if(editId!==null) setPrepList(p=>p.map(m=>m.id===editId?{...m,name:form.name.trim(),unit:form.unit,class:form.class,yieldOverride:form.yieldOverride,ingredients:ings,lastUpdated:now}:m));
    else { const code=genCode("Prep",prepList); setPrepList(p=>[...p,{id:Date.now(),code,name:form.name.trim(),unit:form.unit,class:form.class,yieldOverride:form.yieldOverride,ingredients:ings,lastUpdated:now}]); }
    reset(); showToast(t.savedOk);
  };
  const doEdit=m=>{ setForm({name:m.name,unit:m.unit,class:m.class||"",yieldOverride:m.yieldOverride||"",ingredients:m.ingredients||[]}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setPrepList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };
  const addI=()=>setForm(f=>({...f,ingredients:[...f.ingredients,blank()]}));
  const remI=id=>setForm(f=>({...f,ingredients:f.ingredients.filter(i=>i.id!==id)}));
  const updI=(id,k,v)=>setForm(f=>({...f,ingredients:f.ingredients.map(i=>i.id===id?{...i,[k]:v}:i)}));
  const live=calcPrepCost({ingredients:form.ingredients,unit:form.unit,yieldOverride:form.yieldOverride});

  // Download template — now includes Class column with valid options
  const doDownloadTemplate=()=>{
    const validClasses=cls.join(" / ") || "Sauce / Dough / Mix / Marinade";
    const rows=[
      {"Prep Name":lang==="ar"?"دجاج متبل":"Marinated Chicken","Prep Code":"Prep-00001","Class":cls[0]||"Sauce","Unit":"kg","Ingredient Type":"Raw","Ingredient Code":"Raw-00001","Ingredient Name":lang==="ar"?"دجاج خام":"Raw Chicken","Qty (g/ml)":1000,"Waste %":20},
      {"Prep Name":lang==="ar"?"دجاج متبل":"Marinated Chicken","Prep Code":"Prep-00001","Class":cls[0]||"Sauce","Unit":"kg","Ingredient Type":"Raw","Ingredient Code":"Raw-00002","Ingredient Name":lang==="ar"?"بهارات":"Spices","Qty (g/ml)":30,"Waste %":0},
      {"Prep Name":lang==="ar"?"صوص ثوم":"Garlic Sauce","Prep Code":"Prep-00002","Class":cls[1]||"Sauce","Unit":"kg","Ingredient Type":"Raw","Ingredient Code":"Raw-00003","Ingredient Name":lang==="ar"?"ثوم خام":"Raw Garlic","Qty (g/ml)":500,"Waste %":30},
      {"Prep Name":"","Prep Code":"","Class":`⚡ ${lang==="ar"?"الكلاسات المتاحة":"Valid classes"}: ${validClasses}`,"Unit":"","Ingredient Type":"","Ingredient Code":"","Ingredient Name":"","Qty (g/ml)":"","Waste %":""},
    ];
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=[{wch:22},{wch:14},{wch:16},{wch:10},{wch:16},{wch:14},{wch:25},{wch:12},{wch:10}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Prep Items");
    XLSX.writeFile(wb,"TALA_Prep_Items_Template.xlsx");
  };

  // Export current data
  const doExport=()=>{
    if(!prepList.length){doDownloadTemplate();return;}
    const rows=[];
    prepList.forEach(p=>{
      const ings=p.ingredients||[];
      if(!ings.length){
        rows.push({"Prep Name":p.name,"Prep Code":p.code,"Class":p.class||"","Unit":p.unit||"kg","Ingredient Type":"","Ingredient Code":"","Ingredient Name":"","Qty (g/ml)":"","Waste %":""});
      } else {
        ings.forEach(ing=>{
          let ingCode="",ingName="",ingType="";
          if(ing.source==="prep"){const pp=prepList.find(pp=>String(pp.id)===String(ing.rawId));if(pp){ingCode=pp.code;ingName=pp.name;ingType="Prep";}}
          else{const raw=rawList.find(r=>String(r.id)===String(ing.rawId));if(raw){ingCode=raw.code;ingName=raw.name;ingType="Raw";}}
          rows.push({"Prep Name":p.name,"Prep Code":p.code,"Class":p.class||"","Unit":p.unit||"kg","Ingredient Type":ingType,"Ingredient Code":ingCode,"Ingredient Name":ingName,"Qty (g/ml)":parseFloat(ing.qty)||0,"Waste %":parseFloat(ing.waste)||0});
        });
      }
    });
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=[{wch:22},{wch:14},{wch:16},{wch:10},{wch:16},{wch:14},{wch:25},{wch:12},{wch:10}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Prep Items");
    XLSX.writeFile(wb,"TALA_Prep_Export.xlsx");
  };

  // Smart import — parse → show preview
  const doImport=e=>{
    const file=e.target.files[0]; if(!file) return;
    const r=new FileReader();
    r.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        if(!rows.length){showToast(t.noMatch,"warning");return;}
        const fv=(row,keys)=>{ for(const k of keys){const v=row[k];if(v!==undefined&&v!==null&&String(v).trim()!=="")return String(v).trim();}return ""; };

        // Group by prep name
        const groups={};
        rows.forEach(row=>{
          const pn=fv(row,["Prep Name","اسم Prep","prep name","prep","اسم البريب","البريب","Prep","اسم"]) || fv(row,[Object.keys(row)[0]]);
          const pc=fv(row,["Prep Code","كود Prep","prep code","كود البريب"]);
          const pcls=fv(row,["Class","الكلاس","class","كلاس","Category","فئة"]);
          const punit=fv(row,["Unit","الوحدة","unit","وحدة","UOM"]);
          if(!pn||pn.startsWith("⚡")) return;
          const key=pn.toLowerCase();
          if(!groups[key]) groups[key]={code:pc,name:pn,class:pcls,unit:punit,ingredients:[]};
          else { if(pcls&&!groups[key].class)groups[key].class=pcls; }

          const ingCode=fv(row,["Ingredient Code","كود المكون","code","الكود","Raw Code","Prep Code 2"]);
          const ingName=fv(row,["Ingredient Name","اسم المكون","ingredient","المكون","Raw Name","Prep Name 2","name","الاسم"]);
          const ingTypeRaw=fv(row,["Ingredient Type","نوع المكون","type","النوع","Type"]);
          const ingType=ingTypeRaw.toLowerCase().includes("prep")||ingTypeRaw.includes("بريب")?"prep":"raw";
          const qtyRaw=fv(row,["Qty (g/ml)","الكمية (g/ml)","qty","quantity","الكمية","Qty","weight","الوزن","Amount","g","gram","ml"]);
          let qty=parseFloat(qtyRaw)||0;
          const uomCheck=fv(row,["UOM","وحدة","unit","الوحدة"]).toLowerCase();
          if(["kg","kilo","كيلو","l","liter","litre"].includes(uomCheck) && qty>0 && qty<100) qty*=1000;
          const wasteRaw=fv(row,["Waste %","الهدر %","waste","هدر","Waste","Wastage"]);
          let waste=parseFloat(wasteRaw)||0;
          if(waste>0&&waste<1) waste=Math.round(waste*100*10)/10;
          if(!ingName&&!ingCode) return;
          if(ingType==="prep"){
            const p=prepList.find(p=>p.code===ingCode||p.name.toLowerCase()===ingName.toLowerCase());
            if(p) groups[key].ingredients.push({id:Date.now()+Math.random(),source:"prep",rawId:p.id,qty,waste});
          } else {
            const raw=rawList.find(r=>r.code===ingCode||r.name.toLowerCase()===ingName.toLowerCase());
            if(raw) groups[key].ingredients.push({id:Date.now()+Math.random(),source:"raw",rawId:raw.id,qty,waste});
          }
        });

        // Build preview items (one per prep)
        const items=Object.values(groups).map(pg=>{
          const unitRaw=(pg.unit||"").toLowerCase();
          const unit=["liter","litre","l","لتر"].includes(unitRaw)?"liter":["piece","pcs","حبة"].includes(unitRaw)?"piece":"kg";
          const existing=prepList.find(p=>p.name.toLowerCase()===pg.name.toLowerCase()) ||
                         (pg.code?prepList.find(p=>p.code===pg.code):null);
          return {name:pg.name,code:pg.code,class:pg.class,unit,_ingredients:pg.ingredients,_existing:existing};
        });
        setPreviewData({items,module:"prep"});
        setShowImport(false);
      }catch(err){console.error(err);showToast(lang==="ar"?"خطأ في الملف":"File error","error");}
    };
    r.readAsBinaryString(file); e.target.value="";
  };

  // Commit import after preview
  const confirmImport=(editedItems)=>{
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    let updated=0,added=0;
    setPrepList(prev=>{
      const u=[...prev];
      editedItems.forEach(item=>{
        if(item._skip) return;
        const ings=item._ingredients||[];
        const idx=item._existing?u.findIndex(p=>p.id===item._existing.id):-1;
        if(idx!==-1){
          u[idx]={...u[idx],name:item.name.trim(),class:item.class||u[idx].class,unit:item.unit||u[idx].unit,ingredients:ings,lastUpdated:now};
          updated++;
        } else {
          const newCode=item.code&&!["undefined","nan",""].includes(item.code)?item.code:genCode("Prep",u);
          u.push({id:Date.now()+Math.random(),code:newCode,name:item.name.trim(),class:item.class||"",unit:item.unit||"kg",yieldOverride:"",ingredients:ings,lastUpdated:now});
          added++;
        }
      });
      return u;
    });
    setPreviewData(null);
    showToast(lang==="ar"?`✅ أضيف ${added} بريب | حُدِّث ${updated} بريب`:`✅ Added ${added} | Updated ${updated}`,(added+updated)>0?"success":"warning");
  };

  const [selected,setSelected]=useState(new Set());
  const toggleSel=(id)=>setSelected(p=>{const n=new Set(p);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>setSelected(p=>p.size===filtered.length?new Set():new Set(filtered.map(m=>m.id)));
  const bulkDelete=()=>{setPrepList(p=>p.filter(m=>!selected.has(m.id)));setSelected(new Set());showToast(lang==="ar"?`تم حذف ${selected.size} بريب`:`Deleted ${selected.size} items`,"error");};

  const filtered=prepList.filter(m=>(m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase()))&&(fcls==="all"||m.class===fcls));

  return (
    <div>
      {/* Duplicate warning banner */}
      {dups.length>0&&<div style={{background:"#3d2205",border:"1px solid #ca8a04",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:16}}>⚠️</span>
          <span style={{fontSize:12,color:"#fcd34d",fontWeight:700}}>{lang==="ar"?`يوجد ${dups.length} بريب بأسماء مكررة`:`${dups.length} duplicate prep names found`}</span>
        </div>
        <button onClick={()=>setBulkDelModal({items:dups,module:"prep"})} style={{background:"#ca8a04",color:"#060810",border:"none",borderRadius:6,padding:"5px 12px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
          {lang==="ar"?"عرض وإدارة التكرارات":"View & Manage Duplicates"}
        </button>
      </div>}
      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:240}} placeholder={t.search} value={search} onChange={e=>{setSearch(e.target.value);setSelected(new Set());}}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={()=>setShowImport(true)}>📥 {t.importXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>

      {/* Bulk action bar */}
      {selected.size>0&&hasPerm(mod,"delete")&&<div style={{background:"#1a0f0f",border:"1px solid #dc262666",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <span style={{fontSize:13,color:"#fca5a5",fontWeight:700}}>{lang==="ar"?`تم تحديد ${selected.size} بريب`:`${selected.size} item(s) selected`}</span>
        <div style={{display:"flex",gap:7}}>
          <button onClick={()=>setSelected(new Set())} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:6,padding:"5px 12px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{lang==="ar"?"إلغاء":"Deselect"}</button>
          <button onClick={bulkDelete} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"5px 14px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>🗑 {lang==="ar"?`حذف ${selected.size}`:`Delete ${selected.size}`}</button>
        </div>
      </div>}

      <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}}>
        <button className={`fbtn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all} ({prepList.length})</button>
        {cls.map(c=><button key={c} className={`fbtn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c} ({prepList.filter(m=>m.class===c).length})</button>)}
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>
            {hasPerm(mod,"delete")&&<th style={{width:36,padding:"10px 8px"}}><input type="checkbox" checked={filtered.length>0&&selected.size===filtered.length} onChange={toggleAll} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></th>}
            {["#",t.code,t.name,t.class,t.unit,lang==="ar"?"مكونات":"Ing.",t.yieldWeight,t.costPerUnit,t.usedInProducts,t.actions].map((h,i)=><th key={i}>{h}</th>)}
          </tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={hasPerm(mod,"delete")?11:10} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :filtered.map((m,i)=>{
              const {costPerUnit,yieldKg}=calcPrepCost(m);
              const nProd=prodList.filter(p=>p.ingredients?.some(i=>i.source==="prep"&&String(i.srcId)===String(m.id))).length;
              const isDup=dups.some(d=>d.id===m.id);
              const isSel=selected.has(m.id);
              return <tr key={m.id} style={{background:isSel?C.accent+"12":isDup?"#3d220511":""}}>
                {hasPerm(mod,"delete")&&<td style={{padding:"10px 8px",textAlign:"center"}}><input type="checkbox" checked={isSel} onChange={()=>toggleSel(m.id)} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></td>}
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>
                  {isDup&&<span title={lang==="ar"?"اسم مكرر":"Duplicate"} style={{color:"#f59e0b",marginLeft:4,marginRight:4,fontSize:12}}>⚠</span>}
                  {m.name}
                </td>
                <td><span className="badge badge-cls">{m.class||"—"}</span></td>
                <td><span className={`badge badge-${m.unit}`}>{unitLbl(m.unit,t)}</span></td>
                <td style={{color:C.muted}}>{m.ingredients?.length||0}</td>
                <td style={{color:C.muted}}>{yieldKg.toFixed(3)}</td>
                <td style={{color:C.accent,fontWeight:700}}>{costPerUnit.toFixed(4)}</td>
                <td>{nProd>0?<button onClick={()=>setUsageModal({item:m,type:"product",list:prodList.filter(p=>p.ingredients?.some(i=>i.source==="prep"&&String(i.srcId)===String(m.id)))})} style={{background:"#0a2a1a",color:"#4ade80",border:"1px solid #16a34a33",borderRadius:6,padding:"3px 10px",cursor:"pointer",fontWeight:700,fontSize:13}}>{nProd}</button>:<span style={{color:C.muted}}>0</span>}</td>
                <td><div style={{display:"flex",gap:5}}>
                  <button style={{background:"#0a2a3a",color:"#38bdf8",padding:"4px 10px",fontSize:12,border:"1px solid #0ea5e933",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontWeight:600}} onClick={()=>setViewItem(m)}>{t.view}</button>
                  {hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>doEdit(m)}>{t.edit}</button>}
                  {hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>setDelId(m.id)}>{t.delete}</button>}
                </div></td>
              </tr>;
            })}
          </tbody>
        </table>
      </div></div>

      {viewItem&&(()=>{
        const {costPerUnit,yieldKg,totalCost}=calcPrepCost(viewItem);
        return <div className="overlay" onClick={e=>e.target===e.currentTarget&&setViewItem(null)}>
          <div className="modal modal-lg" style={{maxWidth:820}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
              <div><div style={{fontWeight:800,fontSize:16,color:C.accent}}>{viewItem.name}</div><div style={{fontSize:12,color:C.muted,marginTop:2}}>{viewItem.code} · {viewItem.class||"—"} · {unitLbl(viewItem.unit,t)}</div></div>
              <button onClick={()=>setViewItem(null)} style={{background:"transparent",border:"none",color:C.muted,fontSize:22,cursor:"pointer"}}>✕</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:18}}>
              {[
                {l:lang==="ar"?"عدد المكونات":"Ingredients",v:viewItem.ingredients?.length||0,c:C.blue},
                {l:lang==="ar"?"وزن الناتج":"Yield",v:yieldKg.toFixed(3)+" "+unitLbl(viewItem.unit,t),c:C.green},
                {l:lang==="ar"?"تكلفة الباتش":"Batch Cost",v:(totalCost).toFixed(2),c:"#f87171"},
                {l:lang==="ar"?"تكلفة/وحدة":"Cost/Unit",v:costPerUnit.toFixed(4),c:C.accent},
              ].map((s,i)=><div key={i} style={{background:C.surface,borderRadius:10,padding:"12px 14px",border:`1px solid ${C.border}`}}><div style={{fontSize:10,color:C.muted,marginBottom:4,textTransform:"uppercase"}}>{s.l}</div><div style={{fontSize:17,fontWeight:800,color:s.c}}>{s.v}</div></div>)}
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr>{["#",lang==="ar"?"المادة":"Material",lang==="ar"?"الكمية الخام":"Raw Qty",lang==="ar"?"الهدر %":"Waste %",lang==="ar"?"الكمية الصافية":"Net Qty",lang==="ar"?"سعر/وحدة":"Price/Unit",lang==="ar"?"التكلفة":"Cost"].map((h,i)=><th key={i} style={{padding:"9px 12px",textAlign:lang==="ar"?"right":"left",fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {(viewItem.ingredients||[]).map((ing,i)=>{
                    const raw=rawList.find(r=>String(r.id)===String(ing.rawId)); if(!raw) return null;
                    const qty=parseFloat(ing.qty)||0; const waste=(parseFloat(ing.waste)||0)/100;
                    const netQty=qty*(1-waste); const ingCost=(raw.unit==="piece"?qty:qty/1000)*raw.price;
                    const unit=raw.unit==="kg"?"g":raw.unit==="liter"?"ml":"pcs";
                    return <tr key={i} style={{borderBottom:`1px solid ${C.border}22`}}>
                      <td style={{padding:"10px 12px",color:C.muted,fontSize:11}}>{i+1}</td>
                      <td style={{padding:"10px 12px",fontWeight:600}}>{raw.name} <span style={{color:C.muted,fontSize:11}}>({raw.code})</span></td>
                      <td style={{padding:"10px 12px"}}>{qty.toFixed(0)} {unit}</td>
                      <td style={{padding:"10px 12px"}}>{ing.waste>0?<span style={{background:C.yellow+"22",color:C.yellow,padding:"2px 8px",borderRadius:20,fontSize:12,fontWeight:700}}>{ing.waste}%</span>:<span style={{color:C.muted}}>0%</span>}</td>
                      <td style={{padding:"10px 12px",color:C.green,fontWeight:600}}>{netQty.toFixed(0)} {unit}</td>
                      <td style={{padding:"10px 12px",color:C.muted,fontSize:11}}>{raw.price.toFixed(2)}/{unitLbl(raw.unit,t)}</td>
                      <td style={{padding:"10px 12px",color:C.accent,fontWeight:700}}>{ingCost.toFixed(4)}</td>
                    </tr>;
                  })}
                  <tr style={{background:C.surface,borderTop:`2px solid ${C.border}`}}>
                    <td colSpan={2} style={{padding:"11px 12px",fontWeight:800}}>{lang==="ar"?"الإجمالي":"Total"}</td>
                    <td style={{padding:"11px 12px",fontWeight:700}}>{((viewItem.ingredients||[]).reduce((a,i)=>a+(parseFloat(i.qty)||0),0)).toFixed(0)} g</td>
                    <td/><td style={{padding:"11px 12px",fontWeight:700,color:C.green}}>{((viewItem.ingredients||[]).reduce((a,ing)=>{const w=(parseFloat(ing.waste)||0)/100,q=parseFloat(ing.qty)||0;return a+q*(1-w);},0)).toFixed(0)} g</td>
                    <td/><td style={{padding:"11px 12px",fontWeight:800,color:"#f87171",fontSize:14}}>{(totalCost).toFixed(4)}</td>
                  </tr>
                  <tr style={{background:C.accent+"12"}}>
                    <td colSpan={6} style={{padding:"10px 12px",fontWeight:700,textAlign:lang==="ar"?"right":"left"}}>{lang==="ar"?"تكلفة الكيلو":"Cost per "+unitLbl(viewItem.unit,t)}</td>
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


function ProductsTab({t,lang,C=DARK,prodList,setProdList,rawList,prepList,classes,calcPrepCost,calcProductCost,showToast,hasPerm,mod}) {
  const [searchRaw,setSearchRaw]=useState(""); const [search,setSearch]=useState("");
  const [fcls,setFcls]=useState("all");
  const [showForm,setShowForm]=useState(false); const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",class:"",sellingPrice:"",posSellPrice:"",aggSellPrice:"",stdCost:"",ingredients:[]}); const [errs,setErrs]=useState({});
  const [delId,setDelId]=useState(null); const [showImport,setShowImport]=useState(false); const fileRef=useRef();
  const [ingSearch,setIngSearch]=useState("");
  const [pageSize,setPageSize]=useState(20); const [showCount,setShowCount]=useState(20);
  // Use products classes if available, fallback to raw+prep
  const cls=[...(classes.products||[]),...(classes.raw||[]),...(classes.prep||[])].filter((v,i,a)=>a.indexOf(v)===i);
  const blank=()=>({id:Date.now()+Math.random(),source:"raw",srcId:"",qty:"",waste:"0"});
  const reset=()=>{ setForm({name:"",class:"",sellingPrice:"",posSellPrice:"",aggSellPrice:"",stdCost:"",ingredients:[]}); setErrs({}); setShowForm(false); setEditId(null); };
  const ok=()=>{ const e={}; if(!form.name.trim())e.name=t.required; setErrs(e); return !Object.keys(e).length; };
  const save=()=>{
    if(!ok()) return;
    const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US");
    const ings=form.ingredients.filter(i=>i.srcId&&parseFloat(i.qty)>0);
    const posP=parseFloat(form.posSellPrice)||0;
    const aggP=parseFloat(form.aggSellPrice)||0;
    const sp=parseFloat(form.sellingPrice)||posP||aggP;
    const prod={name:form.name.trim(),class:form.class,sellingPrice:sp,posSellPrice:posP,aggSellPrice:aggP,stdCost:parseFloat(form.stdCost)||0,ingredients:ings,lastUpdated:now};
    if(editId!==null) setProdList(p=>p.map(m=>m.id===editId?{...m,...prod}:m));
    else { const code=genCode("Prod",prodList); setProdList(p=>[...p,{id:Date.now(),code,...prod}]); }
    reset(); showToast(t.savedOk);
  };
  const doEdit=m=>{ setForm({name:m.name,class:m.class||"",sellingPrice:String(m.sellingPrice||""),posSellPrice:String(m.posSellPrice||""),aggSellPrice:String(m.aggSellPrice||""),stdCost:String(m.stdCost||""),ingredients:m.ingredients||[]}); setEditId(m.id); setShowForm(true); };
  const doDelete=id=>{ setProdList(p=>p.filter(m=>m.id!==id)); setDelId(null); showToast(t.deletedOk,"error"); };
  const addI=()=>setForm(f=>({...f,ingredients:[...f.ingredients,blank()]}));
  const remI=id=>setForm(f=>({...f,ingredients:f.ingredients.filter(i=>i.id!==id)}));
  const updI=(id,k,v)=>setForm(f=>({...f,ingredients:f.ingredients.map(i=>i.id===id?{...i,[k]:v}:i)}));
  const live=calcProductCost({ingredients:form.ingredients,sellingPrice:form.sellingPrice,posSellPrice:form.posSellPrice,aggSellPrice:form.aggSellPrice});
  const srcOpts=s=>s==="raw"?rawList:prepList;

  // Debounce
  useEffect(()=>{const timer=setTimeout(()=>setSearch(searchRaw),200);return()=>clearTimeout(timer);},[searchRaw]);
  useEffect(()=>setShowCount(pageSize),[search,fcls,pageSize]);

  const doDownloadTemplate=()=>{
    const rows=[{Code:"Prod-00001",Name:lang==="ar"?"بيتزا مرغريتا":"Margherita Pizza",Class:"Food Item","POS Price":25,"AGG Price":28,"Std Cost":10},{Code:"",Name:"",Class:"","POS Price":"","AGG Price":"","Std Cost":""}];
    const ws=XLSX.utils.json_to_sheet(rows);ws["!cols"]=[{wch:14},{wch:25},{wch:14},{wch:12},{wch:12},{wch:12}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Products");XLSX.writeFile(wb,"TALA_Products_Template.xlsx");
  };
  const doExport=()=>{
    if(!prodList.length){doDownloadTemplate();return;}
    const rows=prodList.map(p=>{const {totalCost,posMargin,aggMargin}=calcProductCost(p);return {Code:p.code,Name:p.name,Class:p.class||"","POS Price":p.posSellPrice||p.sellingPrice||0,"AGG Price":p.aggSellPrice||p.sellingPrice||0,"Std Cost":p.stdCost||0,"Total Cost":totalCost.toFixed(4),"POS Margin %":posMargin.toFixed(2)+"%","AGG Margin %":aggMargin.toFixed(2)+"%"};});
    const ws=XLSX.utils.json_to_sheet(rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Products");XLSX.writeFile(wb,"TALA_Products_Export.xlsx");
  };
  const doImport=e=>{ const file=e.target.files[0]; if(!file) return; const r=new FileReader(); r.onload=ev=>{ try{ const wb=XLSX.read(ev.target.result,{type:"binary"}); const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]); let n=0; const now=new Date().toLocaleDateString(lang==="ar"?"ar-EG":"en-US"); setProdList(prev=>{ const u=[...prev]; rows.forEach(row=>{ const rn=String(row[t.name]||row.Name||row["الاسم"]||"").trim(); const rc=String(row[t.code]||row.Code||row["الكود"]||"").trim(); const rsp=parseFloat(row["Selling Price"]||row["POS Price"]||0); const ragg=parseFloat(row["AGG Price"]||0); let i=u.findIndex(m=>m.name.toLowerCase()===rn.toLowerCase()); if(i===-1) i=u.findIndex(m=>m.code===rc); if(i!==-1){u[i]={...u[i],lastUpdated:now}; if(rsp>0)u[i].posSellPrice=rsp; if(ragg>0)u[i].aggSellPrice=ragg; n++;} }); return u; }); showToast(n>0?`${t.importedOk} ${n} ${t.importedItems}`:t.noMatch,n>0?"success":"warning"); }catch{ showToast("error","error"); } }; r.readAsBinaryString(file); e.target.value=""; };
  const [selected,setSelected]=useState(new Set());
  const toggleSel=(id)=>setSelected(p=>{const n=new Set(p);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleAll=()=>setSelected(p=>p.size===displayed.length?new Set():new Set(displayed.map(m=>m.id)));
  const bulkDelete=()=>{setProdList(p=>p.filter(m=>!selected.has(m.id)));setSelected(new Set());showToast(lang==="ar"?`تم حذف ${selected.size} منتج`:`Deleted ${selected.size} items`,"error");};

  const filtered=prodList.filter(m=>(m.name.toLowerCase().includes(search.toLowerCase())||m.code?.toLowerCase().includes(search.toLowerCase()))&&(fcls==="all"||m.class===fcls));
  const displayed=filtered.slice(0,showCount);
  const hasMore=filtered.length>showCount;
  return (
    <div>
      <div className="card" style={{padding:"12px 14px",marginBottom:12,display:"flex",flexWrap:"wrap",gap:8,alignItems:"center",justifyContent:"space-between"}}>
        <input style={{maxWidth:220}} placeholder={t.search} value={searchRaw} onChange={e=>{setSearchRaw(e.target.value);setSelected(new Set());}}/>
        <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{display:"flex",alignItems:"center",gap:5,background:C.surface,borderRadius:7,padding:"4px 8px",border:`1px solid ${C.border}`}}>
            <span style={{fontSize:11,color:C.muted}}>{t.pageSize}:</span>
            {[5,10,20].map(n=><button key={n} onClick={()=>{setPageSize(n);setShowCount(n);}} style={{background:pageSize===n?C.accent:"transparent",color:pageSize===n?"#080b14":C.muted,border:"none",borderRadius:5,padding:"3px 8px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{n}</button>)}
          </div>
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={doExport}>{t.exportXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-secondary" onClick={()=>setShowImport(true)}>📥 {t.importXlsx}</button>}
          {hasPerm(mod,"edit")&&<button className="btn btn-primary" onClick={()=>{reset();setShowForm(true);}}>+ {t.add}</button>}
        </div>
      </div>

      {/* Bulk action bar */}
      {selected.size>0&&hasPerm(mod,"delete")&&<div style={{background:"#1a0f0f",border:"1px solid #dc262666",borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
        <span style={{fontSize:13,color:"#fca5a5",fontWeight:700}}>{lang==="ar"?`تم تحديد ${selected.size} منتج`:`${selected.size} item(s) selected`}</span>
        <div style={{display:"flex",gap:7}}>
          <button onClick={()=>setSelected(new Set())} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:6,padding:"5px 12px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>{lang==="ar"?"إلغاء":"Deselect"}</button>
          <button onClick={bulkDelete} style={{background:"#2a0f0f",color:"#f87171",border:"1px solid #dc262633",borderRadius:6,padding:"5px 14px",fontSize:12,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>🗑 {lang==="ar"?`حذف ${selected.size}`:`Delete ${selected.size}`}</button>
        </div>
      </div>}

      <div style={{display:"flex",gap:7,marginBottom:10,flexWrap:"wrap"}}>
        <button className={`filter-btn${fcls==="all"?" active":""}`} onClick={()=>setFcls("all")}>{t.all} ({prodList.length})</button>
        {cls.map(c=><button key={c} className={`filter-btn${fcls===c?" active":""}`} onClick={()=>setFcls(c)}>{c} ({prodList.filter(m=>m.class===c).length})</button>)}
      </div>
      <div className="card" style={{overflow:"hidden"}}><div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>
            {hasPerm(mod,"delete")&&<th style={{width:36,padding:"10px 8px"}}><input type="checkbox" checked={displayed.length>0&&selected.size===displayed.length} onChange={toggleAll} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></th>}
            {["#",t.code,t.productName,t.class,lang==="ar"?"مكونات":"Ing.",t.totalCost,t.stdCost,lang==="ar"?"POS":"POS",lang==="ar"?"هامش POS":"POS%",lang==="ar"?"AGG":"AGG",lang==="ar"?"هامش AGG":"AGG%",(hasPerm(mod,"edit")||hasPerm(mod,"delete"))?t.actions:""].filter(Boolean).map((h,i)=><th key={i}>{h}</th>)}
          </tr></thead>
          <tbody>
            {displayed.length===0?<tr><td colSpan={12} style={{textAlign:"center",padding:"40px",color:C.muted}}>{t.noData}</td></tr>
            :displayed.map((m,i)=>{ const {totalCost,posMargin,aggMargin}=calcProductCost(m); const isSel=selected.has(m.id);
              const posP=parseFloat(m.posSellPrice||m.sellingPrice||0); const aggP=parseFloat(m.aggSellPrice||m.sellingPrice||0);
              return(
              <tr key={m.id} style={{background:isSel?C.accent+"12":""}}>
                {hasPerm(mod,"delete")&&<td style={{padding:"10px 8px",textAlign:"center"}}><input type="checkbox" checked={isSel} onChange={()=>toggleSel(m.id)} style={{width:14,height:14,accentColor:C.accent,cursor:"pointer"}}/></td>}
                <td style={{color:C.muted,fontSize:11}}>{i+1}</td>
                <td><code style={{background:C.surface,padding:"2px 6px",borderRadius:4,fontSize:11,color:"#94a3b8"}}>{m.code}</code></td>
                <td style={{fontWeight:600}}>{m.name}</td>
                <td><span className="badge badge-cls">{m.class||"—"}</span></td>
                <td style={{color:C.muted}}>{m.ingredients?.length||0}</td>
                <td style={{color:"#f87171",fontWeight:600}}>{totalCost.toFixed(2)}</td>
                <td style={{color:DARK.muted,fontWeight:500}}>{m.stdCost>0?m.stdCost.toFixed(2):"—"}</td>
                <td style={{color:"#4ade80",fontWeight:600}}>{posP>0?posP.toFixed(2):"—"}</td>
                <td><span style={{color:posMargin>30?"#4ade80":posMargin>15?"#fbbf24":"#f87171",fontWeight:700}}>{posP>0?posMargin.toFixed(1)+"%":"—"}</span></td>
                <td style={{color:"#60a5fa",fontWeight:600}}>{aggP>0?aggP.toFixed(2):"—"}</td>
                <td><span style={{color:aggMargin>30?"#4ade80":aggMargin>15?"#fbbf24":"#f87171",fontWeight:700}}>{aggP>0?aggMargin.toFixed(1)+"%":"—"}</span></td>
                {(hasPerm(mod,"edit")||hasPerm(mod,"delete"))&&<td><div style={{display:"flex",gap:5}}>{hasPerm(mod,"edit")&&<button className="btn-sm-e" onClick={()=>doEdit(m)}>{t.edit}</button>}{hasPerm(mod,"delete")&&<button className="btn-sm-d" onClick={()=>setDelId(m.id)}>{t.delete}</button>}</div></td>}
              </tr>
            );})}
          </tbody>
        </table>
      </div></div>
      {/* Pagination */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginTop:10}}>
        <span style={{fontSize:12,color:C.muted}}>{lang==="ar"?`عرض ${displayed.length} من ${filtered.length}`:`Showing ${displayed.length} of ${filtered.length}`}</span>
        {hasMore&&<button onClick={()=>setShowCount(c=>c+pageSize)} style={{background:C.surface,color:C.accent,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{t.showMore} (+{Math.min(pageSize,filtered.length-showCount)})</button>}
        {showCount>pageSize&&!hasMore&&<button onClick={()=>setShowCount(pageSize)} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 16px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{t.showLess}</button>}
      </div>
      {showForm&&<div className="overlay" onClick={e=>e.target===e.currentTarget&&reset()}><div className="modal modal-lg">
        <h2 style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:16}}>{editId?t.edit:t.add} — {t.products}</h2>
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:10,marginBottom:10}}>
          <div><label className="lbl">{t.productName}</label><input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>{errs.name&&<div className="err">{errs.name}</div>}</div>
          <div><label className="lbl">{t.class}</label><select value={form.class} onChange={e=>setForm({...form,class:e.target.value})}><option value="">—</option>{cls.map(c=><option key={c} value={c}>{c}</option>)}</select></div>
        </div>
        {/* POS / AGG pricing grid */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10,background:C.surface,borderRadius:9,padding:"12px 14px",border:`1px solid ${C.border}`}}>
          <div>
            <label className="lbl" style={{color:"#22c55e"}}>🏪 {t.posSellPrice}</label>
            <input type="number" min="0" step="0.01" value={form.posSellPrice} onChange={e=>setForm({...form,posSellPrice:e.target.value})} placeholder="0.00" style={{borderColor:form.posSellPrice?"#22c55e44":undefined}}/>
            {parseFloat(form.posSellPrice)>0&&<div style={{fontSize:11,color:"#22c55e",marginTop:3}}>هامش: {live.posMargin.toFixed(1)}%</div>}
          </div>
          <div>
            <label className="lbl" style={{color:"#3b82f6"}}>🛵 {t.aggSellPrice}</label>
            <input type="number" min="0" step="0.01" value={form.aggSellPrice} onChange={e=>setForm({...form,aggSellPrice:e.target.value})} placeholder="0.00" style={{borderColor:form.aggSellPrice?"#3b82f644":undefined}}/>
            {parseFloat(form.aggSellPrice)>0&&<div style={{fontSize:11,color:"#3b82f6",marginTop:3}}>هامش: {live.aggMargin.toFixed(1)}%</div>}
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
        {form.ingredients.map(ing=>(
          <IngRowProd key={ing.id} ing={ing} rawList={rawList} prepList={prepList} lang={lang} t={t} C={C}
            onUpdate={(field,val)=>updI(ing.id,field,val)}
            onRemove={()=>remI(ing.id)}
          />
        ))}
        <div style={{background:C.bg,borderRadius:8,padding:"10px 13px",marginTop:8,display:"flex",gap:16,flexWrap:"wrap"}}>
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

// ═══════════════════════════════════════════════════════════════
// CLASSES
// ═══════════════════════════════════════════════════════════════
function ClassesTab({t,lang,C=DARK,classes,setClasses,showToast,hasPerm,mod}) {
  const [et,setEt]=useState(null); const [ei,setEi]=useState(null); const [val,setVal]=useState(""); const [err,setErr]=useState("");
  const secs=[
    {key:"raw",label:t.rawMat,cat:t.rawMatCategory},
    {key:"prep",label:t.prepItem,cat:t.prepCategory},
    {key:"products",label:t.products,cat:lang==="ar"?"كلاسات المنتجات":"Product Classes"},
  ];
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
