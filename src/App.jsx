import { useState, useEffect } from "react";

const STORAGE_KEYS = { products: "kdp-products", movements: "kdp-movements" };

// Limpia datos de ejemplo y normaliza productIds
if (typeof localStorage !== "undefined") {
  const flagKey = "kdp-cleared-v2";
  if (!localStorage.getItem(flagKey)) {
    localStorage.removeItem("kdp-products");
    localStorage.removeItem("kdp-movements");
    localStorage.setItem(flagKey, "1");
  }
  // Normaliza: si solo hay un producto, todos los movimientos usan su id
  try {
    const prods = JSON.parse(localStorage.getItem("kdp-products") || "[]");
    const movs = JSON.parse(localStorage.getItem("kdp-movements") || "[]");
    if (prods.length === 1 && movs.length > 0) {
      const pid = prods[0].id;
      const fixed = movs.map(m => ({...m, productId: pid}));
      localStorage.setItem("kdp-movements", JSON.stringify(fixed));
    }
  } catch(e) {}
}

const EMOJIS = ["📚","🥗","🍳","🏋️","✍️","🎯","💡","🌿","🔥","⭐"];
const COLORS = ["#e8f5e9","#e3f2fd","#fce4ec","#fff8e1","#f3e5f5","#e0f7fa","#fbe9e7","#f1f8e9"];

const defaultProducts = [];
const defaultMovements = [];

const CATS_GASTO = [
  "Amazon ADS","Claude","Canva","Publi TikTok","Pago reseñadores",
  "Sequra-Mastermind","Master Publisher","App web","Web Empresa","Compras libros",
  "Plantillas KDP","Herramientas","Otros"
];

const MONTHS_ES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

function fmtAbs(n) { return n.toFixed(2).replace(".", ",") + " €"; }
function fmtSigned(n) { return (n >= 0 ? "+" : "-") + fmtAbs(Math.abs(n)); }
async function exportToXLSX(data, filename, periodo) {
  // Load SheetJS from CDN
  if (!window.XLSX) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
  }
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();

  // Build rows: title + headers + data
  const titleRow = [periodo];
  const headers = ["Fecha", "Concepto", "Importe (€)", "Mes devengo", "Notas"];
  const rows = data.map(m => [
    m.date,
    m.concept,
    parseFloat(m.amount.toFixed(2)),
    m.devengoMonth || "",
    m.notes || ""
  ]);

  const wsData = [titleRow, [], headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Style title: merge cells and bold
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];
  if (ws["A1"]) {
    ws["A1"].s = {
      font: { bold: true, sz: 16 },
      alignment: { horizontal: "center" }
    };
  }
  // Style headers row (row index 2)
  headers.forEach((_, i) => {
    const cellRef = XLSX.utils.encode_cell({ r: 2, c: i });
    if (ws[cellRef]) ws[cellRef].s = { font: { bold: true } };
  });

  // Column widths
  ws["!cols"] = [{ wch: 14 }, { wch: 35 }, { wch: 14 }, { wch: 14 }, { wch: 30 }];

  XLSX.utils.book_append_sheet(wb, ws, "Datos");
  XLSX.writeFile(wb, filename);
}

function monthName(ym) {
  const [y, m] = ym.split("-");
  return MONTHS_ES[parseInt(m)-1] + " De " + y;
}

function useStorage(key, def) {
  const [val, setVal] = useState(() => {
    try { const s = localStorage.getItem(key); return s ? JSON.parse(s) : def; } catch { return def; }
  });
  useEffect(() => { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} }, [key, val]);
  return [val, setVal];
}

function calcStats(movements, pid, prefix, ignorePid = false) {
  const mvs = movements.filter(m => (ignorePid || m.productId === pid) && (!prefix || m.date.startsWith(prefix)));
  const ingresos = mvs.filter(m => m.type === "venta").reduce((a, m) => a + m.amount, 0);
  const gastos   = mvs.filter(m => m.type === "gasto").reduce((a, m) => a + m.amount, 0);
  const resultado = ingresos - gastos;
  const roi = gastos > 0 ? (resultado / gastos) * 100 : (ingresos > 0 ? 100 : 0);
  return { ingresos, gastos, resultado, roi };
}

// ── COMPONENTS ──────────────────────────────────────────────────────────────

function MovementRow({ m, onDelete, onEdit, showProduct, products }) {
  const prod = products.find(p => p.id === m.productId);
  return (
    <div style={{ background: "#fff", borderRadius: 14, padding: "13px 16px", marginBottom: 8,
      boxShadow: "0 1px 4px rgba(0,0,0,0.05)", display: "flex", alignItems: "center", gap: 12 }}>
      <div style={{ width: 38, height: 38, borderRadius: 10, flexShrink: 0,
        background: m.type === "venta" ? "#e8f5e9" : "#fdecea",
        display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17 }}>
        {m.type === "venta" ? "📈" : "📉"}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#1a1a1a",
          whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{m.concept}</p>
        <p style={{ margin: "2px 0 0", fontSize: 11, color: "#aaa" }}>
          {showProduct ? prod?.name?.slice(0,16)+"... · " : "Compras · "}{m.date}
          {m.devengoMonth && <span style={{ marginLeft: 5, background: "#fff8e1", color: "#b8860b", borderRadius: 4, padding: "1px 5px", fontWeight: 600 }}>📅 {m.devengoMonth}</span>}
        </p>
      </div>
      <span style={{ fontSize: 14, fontWeight: 700, fontFamily: "'DM Mono', monospace", whiteSpace: "nowrap",
        color: m.type === "venta" ? "#1a7a4a" : "#c0392b" }}>
        {m.type === "venta" ? "+" : "-"}{fmtAbs(m.amount)}
      </span>
      <button onClick={() => onEdit && onEdit(m)}
        style={{ background: "none", border: "none", color: "#bbb", fontSize: 15, cursor: "pointer", padding: "0 2px", flexShrink: 0 }}>✏️</button>
      <button onClick={() => onDelete(m.id)}
        style={{ background: "none", border: "none", color: "#ddd", fontSize: 15, cursor: "pointer", padding: "0 2px", flexShrink: 0 }}>🗑</button>
    </div>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex",
      alignItems: "flex-end", zIndex: 200 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: "22px 22px 0 0", padding: "24px 20px 44px",
        width: "100%", maxWidth: 420, margin: "0 auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700 }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: "#aaa" }}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ fontSize: 11, color: "#aaa", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, display: "block", marginBottom: 5 }}>{label}</label>
      {children}
    </div>
  );
}

const inputStyle = { width: "100%", border: "1.5px solid #e8e8e8", borderRadius: 11, padding: "11px 13px",
  fontSize: 14, outline: "none", background: "#fafafa", boxSizing: "border-box", fontFamily: "inherit" };

// ── MAIN APP ─────────────────────────────────────────────────────────────────

export default function App() {
  const [products,  setProducts]  = useStorage(STORAGE_KEYS.products,  defaultProducts);
  const [movements, setMovements] = useStorage(STORAGE_KEYS.movements, defaultMovements);
  const [tab,       setTab]       = useState("inicio");
  const [activePid, setActivePid] = useState(() => {
    try {
      const s = localStorage.getItem("kdp-products");
      const prods = s ? JSON.parse(s) : [];
      return prods.length > 0 ? prods[0].id : 1;
    } catch { return 1; }
  });
  const [modal,     setModal]     = useState(null);
  const [form,      setForm]      = useState({});
  const [editingId, setEditingId] = useState(null);
  const [selectedMonth, setSelectedMonth] = useState(null);
  const [filterMonth, setFilterMonth] = useState("");
  const [filterType, setFilterType] = useState("all"); // gastos/ventas filter
  const [informeYear, setInformeYear] = useState(new Date().getFullYear());

  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`;
  const activeProduct = products.find(p => p.id === activePid) || products[0];

  function setF(k, v) { setForm(f => ({...f, [k]: v})); }

  // ── ACTIONS
  function deleteMovement(id) {
    if (!window.confirm("¿Seguro que quieres borrar este movimiento?")) return;
    setMovements(prev => prev.filter(m => m.id !== id));
  }

  function deleteProduct(pid) {
    if (!window.confirm("¿Seguro que quieres borrar este producto y todos sus movimientos?")) return;
    setProducts(prev => prev.filter(p => p.id !== pid));
    setMovements(prev => prev.filter(m => m.productId !== pid));
    if (activePid === pid) {
      const remaining = products.filter(p => p.id !== pid);
      if (remaining.length > 0) setActivePid(remaining[0].id);
    }
  }

  function openEdit(m) {
    setEditingId(m.id);
    setForm({
      concept: m.concept,
      amount: String(m.amount),
      date: m.date,
      devengoMonth: m.devengoMonth || "",
      notes: m.notes || "",
      units: String(m.units || 1),
    });
    setModal(m.type === "venta" ? "edit-venta" : "edit-gasto");
  }

  function updateMovement() {
    const amt = parseFloat((form.amount || "0").replace(",", "."));
    if (!amt) return;
    setMovements(prev => prev.map(m => m.id === editingId ? {
      ...m,
      concept: form.concept || m.concept,
      amount: amt,
      date: form.date || m.date,
      devengoMonth: form.devengoMonth || null,
      notes: form.notes || null,
    } : m));
    setModal(null); setForm({}); setEditingId(null);
  }

  function saveMovement(type) {
    const amt = parseFloat((form.amount || "0").replace(",", "."));
    if (!amt) return;
    const concept = form.concept === "__custom" ? (form.customConcept || "Gasto") : (form.concept || activeProduct?.name);
    setMovements(prev => [...prev, {
      id: Date.now(), productId: activePid, type, concept, amount: amt,
      date: form.date || now.toISOString().slice(0,10)
    }]);
    setModal(null); setForm({});
  }

  function saveProduct() {
    if (!form.name) return;
    const p = { id: Date.now(), name: form.name, emoji: form.emoji || "📚", color: form.color || "#e8f5e9" };
    setProducts(prev => [...prev, p]);
    setActivePid(p.id);
    setModal(null); setForm({});
  }

  // ── STATS
  const monthStats = calcStats(movements, activePid, currentMonth);
  const allStats   = calcStats(movements, activePid);

  const totalGastos = movements.filter(m => m.type === "gasto").reduce((a, m) => a + m.amount, 0);
  const totalVentas = movements.filter(m => m.type === "venta").reduce((a, m) => a + m.amount, 0);

  // ── INFORMES: months for active product in selected year
  const yearPrefix = String(informeYear);
  const availableYears = [...new Set(movements.filter(m => m.productId === activePid).map(m => m.date.slice(0,4)))].sort().reverse();
  const yearStats = calcStats(movements, activePid, yearPrefix);

  // Meses con cualquier movimiento del producto activo en el año seleccionado
  const monthsWithData = [...new Set(
    movements
      .filter(m => m.productId === activePid && m.date.startsWith(yearPrefix))
      .map(m => m.date.slice(0,7))
  )];
  const monthsInYear = Array.from({length: 12}, (_, i) => {
    const ym = `${informeYear}-${String(i+1).padStart(2,"0")}`;
    return { ym, ...calcStats(movements, activePid, ym) };
  }).filter(m => monthsWithData.includes(m.ym)).reverse();

  const maxBar = Math.max(...monthsInYear.map(m => Math.max(m.ingresos, m.gastos)), 1);

  // product rankings all time
  const rankings = [...products].map(p => ({ p, ...calcStats(movements, p.id) })).sort((a,b) => b.resultado - a.resultado);

  // ── FILTERED lists for gastos/ventas tabs
  const allMovSorted = [...movements].sort((a,b) => b.date.localeCompare(a.date));
  const gastosList = allMovSorted.filter(m => m.productId === activePid && m.type === "gasto" && (!filterMonth || m.date.startsWith(filterMonth)));
  const ventasList = allMovSorted.filter(m => m.productId === activePid && m.type === "venta" && (!filterMonth || m.date.startsWith(filterMonth)));

  const resultado = monthStats.resultado;
  const isPos = resultado >= 0;

  // ── RENDER ────────────────────────────────────────────────────────────────
  return (
    <div style={{ fontFamily: "'DM Sans', sans-serif", background: "#f7f5f0", minHeight: "100vh",
      maxWidth: 420, margin: "0 auto", paddingBottom: 80 }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,400;0,500;0,600;0,700;1,400&family=DM+Mono:wght@500&display=swap" rel="stylesheet" />

      {/* ── INICIO ── */}
      {tab === "inicio" && (<>
        <div style={{ padding: "28px 20px 12px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div>
              <h1 style={{ margin: 0, fontSize: 26, fontWeight: 700, color: "#1a1a1a" }}>Hola 👋</h1>
              <p style={{ margin: "2px 0 0", fontSize: 13, color: "#999" }}>
                Resumen de {MONTHS_ES[now.getMonth()].toLowerCase()} de {now.getFullYear()}
              </p>
            </div>
            <button onClick={() => { setModal("product"); setForm({ emoji: "📚", color: "#e8f5e9" }); }}
              style={{ background: "#1a1a1a", color: "#fff", border: "none", borderRadius: 10,
                padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>+ Libro</button>
          </div>

          {/* Product selector */}
          <div style={{ marginTop: 14, background: "#fff", borderRadius: 14, padding: "12px 14px",
            display: "flex", alignItems: "center", gap: 12, boxShadow: "0 1px 4px rgba(0,0,0,0.06)", overflow: "hidden" }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: activeProduct?.color,
              display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, flexShrink: 0 }}>
              {activeProduct?.emoji}
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <p style={{ margin: 0, fontSize: 10, color: "#bbb", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5 }}>Producto activo</p>
              <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#1a1a1a",
                whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{activeProduct?.name}</p>
            </div>
            <select value={activePid} onChange={e => setActivePid(Number(e.target.value))}
              style={{ border: "none", background: "transparent", fontSize: 16, cursor: "pointer", outline: "none", flexShrink: 0, maxWidth: 28 }}>
              {products.map(p => <option key={p.id} value={p.id}>{p.emoji} {p.name}</option>)}
            </select>
          </div>
        </div>

        <div style={{ padding: "0 20px" }}>
          {/* Resultado */}
          <div style={{ background: isPos ? "#1a7a4a" : "#c0392b", borderRadius: 18, padding: "20px 22px",
            marginBottom: 14, boxShadow: "0 4px 18px rgba(0,0,0,0.13)" }}>
            <p style={{ margin: "0 0 4px", fontSize: 12, color: "rgba(255,255,255,0.7)", fontWeight: 500 }}>
              📊 Resultado final del mes
            </p>
            <p style={{ margin: 0, fontSize: 38, fontWeight: 700, color: "#fff", fontFamily: "'DM Mono', monospace", letterSpacing: -1 }}>
              {resultado >= 0 ? "+" : ""}{fmtAbs(Math.abs(resultado))}
            </p>
            <p style={{ margin: "6px 0 0", fontSize: 12, color: "rgba(255,255,255,0.65)" }}>
              {isPos ? "¡Mes rentable! 🎉" : "Estás gastando más de lo que ingresas"}
            </p>
          </div>

          {/* Ganancias / Gastos */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
            {[["📈","Ganancias", monthStats.ingresos, "#1a7a4a"], ["📉","Gastos", monthStats.gastos, "#c0392b"]].map(([ic,lb,val,col]) => (
              <div key={lb} style={{ background: "#fff", borderRadius: 14, padding: "14px 16px", boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6 }}>
                  <span style={{ fontSize: 15 }}>{ic}</span>
                  <span style={{ fontSize: 12, color: "#aaa", fontWeight: 500 }}>{lb}</span>
                </div>
                <p style={{ margin: 0, fontSize: 20, fontWeight: 700, color: "#1a1a1a", fontFamily: "'DM Mono', monospace" }}>{fmtAbs(val)}</p>
              </div>
            ))}
          </div>

          {/* ROI */}
          <div style={{ background: "#fff", borderRadius: 14, padding: "12px 16px", marginBottom: 14,
            boxShadow: "0 1px 4px rgba(0,0,0,0.06)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 13, color: "#555", fontWeight: 500 }}>ROI del mes</span>
            <span style={{ fontSize: 22, fontWeight: 700, fontFamily: "'DM Mono', monospace",
              color: monthStats.roi >= 0 ? "#1a7a4a" : "#c0392b" }}>
              {monthStats.roi >= 0 ? "+" : ""}{monthStats.roi.toFixed(0)}%
            </span>
          </div>

          {/* Botones */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 24 }}>
            <button onClick={() => { setModal("venta"); setForm({ date: now.toISOString().slice(0,10), concept: activeProduct?.name }); }}
              style={{ background: "#1a7a4a", color: "#fff", border: "none", borderRadius: 14, padding: "15px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
              + Nueva venta
            </button>
            <button onClick={() => { setModal("gasto"); setForm({ date: now.toISOString().slice(0,10) }); }}
              style={{ background: "#c0392b", color: "#fff", border: "none", borderRadius: 14, padding: "15px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
              + Nuevo gasto
            </button>
          </div>

          {/* Productos */}
          <h2 style={{ fontSize: 17, fontWeight: 700, margin: "0 0 10px" }}>Tus productos</h2>
          {products.map(p => {
            const ms = calcStats(movements, p.id, currentMonth);
            const at = calcStats(movements, p.id);
            return (
              <div key={p.id} onClick={() => setActivePid(p.id)}
                style={{ background: "#fff", borderRadius: 14, padding: "13px 15px", marginBottom: 8, cursor: "pointer",
                  boxShadow: "0 1px 4px rgba(0,0,0,0.06)", border: p.id === activePid ? "2px solid #1a1a1a" : "2px solid transparent" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
                  <div style={{ width: 36, height: 36, borderRadius: 9, background: p.color, flexShrink: 0,
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>{p.emoji}</div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#1a1a1a",
                      whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</p>
                    <p style={{ margin: "2px 0 0", fontSize: 11, color: "#bbb" }}>
                      Acumulado: <span style={{ color: at.resultado >= 0 ? "#1a7a4a" : "#c0392b", fontWeight: 600 }}>{fmtSigned(at.resultado)}</span>
                    </p>
                  </div>
                  <div style={{ textAlign: "right", flexShrink: 0 }}>
                    <p style={{ margin: 0, fontSize: 14, fontWeight: 700, fontFamily: "'DM Mono', monospace",
                      color: ms.resultado >= 0 ? "#1a7a4a" : "#c0392b" }}>{fmtSigned(ms.resultado)}</p>
                    <p style={{ margin: "2px 0 0", fontSize: 11, color: "#bbb" }}>ROI {ms.roi.toFixed(0)}%</p>
                  </div>
                </div>
              </div>
            );
          })}

          <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 14px", background: "#fff",
            borderRadius: 12, marginBottom: 20, boxShadow: "0 1px 3px rgba(0,0,0,0.04)" }}>
            <span style={{ fontSize: 13, color: "#555" }}>Beneficio acumulado (productos)</span>
            <span style={{ fontSize: 13, fontWeight: 700, fontFamily: "'DM Mono', monospace",
              color: products.reduce((a,p) => a + calcStats(movements,p.id).resultado, 0) >= 0 ? "#1a7a4a" : "#c0392b" }}>
              {fmtSigned(products.reduce((a,p) => a + calcStats(movements,p.id).resultado, 0))}
            </span>
          </div>

          {/* Últimos movimientos */}
          <h2 style={{ fontSize: 17, fontWeight: 700, margin: "0 0 10px" }}>Últimos movimientos</h2>
          {[...movements].sort((a,b) => b.date.localeCompare(a.date)).slice(0,8).map(m => (
            <MovementRow key={m.id} m={m} onDelete={deleteMovement} onEdit={openEdit} showProduct products={products} />
          ))}
        </div>
      </>)}

      {/* ── PRODUCTOS ── */}
      {tab === "productos" && (
        <div style={{ padding: "28px 20px 0" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
            <h1 style={{ margin: 0, fontSize: 24, fontWeight: 700 }}>Productos</h1>
            <button onClick={() => { setModal("product"); setForm({ emoji: "📚", color: "#e8f5e9" }); }}
              style={{ background: "#1a1a1a", color: "#fff", border: "none", borderRadius: 10,
                padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>+ Nuevo</button>
          </div>
          {products.map(p => {
            const at = calcStats(movements, p.id);
            const ms = calcStats(movements, p.id, currentMonth);
            // Months with data
            const prodMonths = [...new Set(movements.map(m => m.date.slice(0,7)))].sort().reverse();
            return (
              <div key={p.id} onClick={() => setActivePid(p.id)}
                style={{ background: "#fff", borderRadius: 16, padding: "17px", marginBottom: 12, cursor: "pointer",
                  boxShadow: "0 1px 6px rgba(0,0,0,0.07)", border: p.id === activePid ? "2px solid #1a1a1a" : "2px solid transparent" }}>
                <div style={{ display: "flex", gap: 13, alignItems: "center", marginBottom: 13 }}>
                  <div style={{ width: 46, height: 46, borderRadius: 12, background: p.color, flexShrink: 0,
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 23 }}>{p.emoji}</div>
                  <div style={{ flex: 1 }}>
                    <p style={{ margin: 0, fontSize: 14, fontWeight: 700, color: "#1a1a1a" }}>{p.name}</p>
                    {p.id === activePid && <span style={{ fontSize: 10, background: "#1a1a1a", color: "#fff",
                      borderRadius: 5, padding: "2px 7px", fontWeight: 700 }}>ACTIVO</span>}
                  </div>
                  <button onClick={e => { e.stopPropagation(); deleteProduct(p.id); }}
                    style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", color: "#ddd", flexShrink: 0 }}>🗑</button>
                </div>
                {/* Totales acumulados */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 7, marginBottom: 14 }}>
                  {[["Ingresos", fmtAbs(at.ingresos), "#1a7a4a"],["Gastos", fmtAbs(at.gastos), "#c0392b"],
                    ["Resultado", fmtSigned(at.resultado), at.resultado>=0?"#1a7a4a":"#c0392b"],
                    ["ROI", at.roi.toFixed(0)+"%", at.roi>=0?"#1a7a4a":"#c0392b"]].map(([l,v,c]) => (
                    <div key={l} style={{ background: "#f7f5f0", borderRadius: 9, padding: "9px 8px" }}>
                      <p style={{ margin: 0, fontSize: 9, color: "#bbb", fontWeight: 700, textTransform: "uppercase" }}>{l}</p>
                      <p style={{ margin: "4px 0 0", fontSize: 11, fontWeight: 700, color: c, fontFamily: "'DM Mono', monospace" }}>{v}</p>
                    </div>
                  ))}
                </div>
                {/* Desglose por meses */}
                <p style={{ margin: "0 0 8px", fontSize: 11, fontWeight: 700, color: "#bbb", textTransform: "uppercase", letterSpacing: 0.5 }}>Por mes</p>
                {prodMonths.map(ym => {
                  const ms2 = calcStats(movements, p.id, ym);
                  if (ms2.ingresos === 0 && ms2.gastos === 0) return null;
                  return (
                    <div key={ym} style={{ borderTop: "1px solid #f0f0f0", paddingTop: 8, marginBottom: 8 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 5 }}>
                        <span style={{ fontSize: 12, fontWeight: 600, color: "#555" }}>{monthName(ym)}</span>
                        <span style={{ fontSize: 12, fontWeight: 700, fontFamily: "'DM Mono', monospace",
                          color: ms2.resultado>=0?"#1a7a4a":"#c0392b" }}>{fmtSigned(ms2.resultado)}</span>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 5 }}>
                        {[["Ventas", fmtAbs(ms2.ingresos), "#1a7a4a"],["Gastos", fmtAbs(ms2.gastos), "#c0392b"],
                          ["Resultado", fmtSigned(ms2.resultado), ms2.resultado>=0?"#1a7a4a":"#c0392b"],
                          ["ROI", ms2.roi.toFixed(0)+"%", ms2.roi>=0?"#1a7a4a":"#c0392b"]].map(([l,v,c]) => (
                          <div key={l} style={{ background: "#f7f5f0", borderRadius: 7, padding: "6px 6px" }}>
                            <p style={{ margin: 0, fontSize: 8, color: "#bbb", fontWeight: 700, textTransform: "uppercase" }}>{l}</p>
                            <p style={{ margin: "3px 0 0", fontSize: 10, fontWeight: 700, color: c, fontFamily: "'DM Mono', monospace" }}>{v}</p>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      )}

      {/* ── GASTOS ── */}
      {tab === "gastos" && (
        <div style={{ padding: "28px 20px 0" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
            <div>
              <h1 style={{ margin: 0, fontSize: 24, fontWeight: 700 }}>Gastos</h1>
              <p style={{ margin: "3px 0 0", fontSize: 13, color: "#999" }}>Total: {fmtAbs(gastosList.reduce((a,m) => a + m.amount, 0))}</p>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => exportToXLSX(gastosList, "gastos-kdp.xlsx", filterMonth ? `Gastos · ${MONTHS_ES[parseInt(filterMonth.split("-")[1])-1]} ${filterMonth.split("-")[0]}` : `Gastos · Histórico completo ${new Date().getFullYear()}`)}
                style={{ background: "#f0f0f0", color: "#555", border: "none", borderRadius: 22,
                  padding: "10px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>📥 Excel</button>
              <button onClick={() => { setModal("gasto"); setForm({ date: now.toISOString().slice(0,10) }); }}
                style={{ background: "#c0392b", color: "#fff", border: "none", borderRadius: 22,
                  padding: "10px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>+ Nuevo gasto</button>
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, margin: "16px 0 14px", alignItems: "center" }}>
            <div style={{ flex: 1, background: "#fff", borderRadius: 22, padding: "6px 14px",
              fontSize: 13, color: "#888", boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
              {filterMonth ? `Filtrando: ${MONTHS_ES[parseInt(filterMonth.split("-")[1])-1]} ${filterMonth.split("-")[0]}` : "Todo El Histórico"}
            </div>
            <input type="month" value={filterMonth} onChange={e => setFilterMonth(e.target.value)}
              style={{ border: "1.5px solid #e8e8e8", borderRadius: 11, padding: "6px 12px", fontSize: 12,
                background: "#fff", outline: "none", fontFamily: "inherit", cursor: "pointer" }} />
            {filterMonth && <button onClick={() => setFilterMonth("")}
              style={{ background: "#f0f0f0", border: "none", borderRadius: 11, padding: "6px 10px", fontSize: 12, cursor: "pointer" }}>✕</button>}
          </div>

          {gastosList.length === 0 && (
            <div style={{ textAlign: "center", padding: "50px 0", color: "#ccc" }}>
              <p style={{ fontSize: 36 }}>💸</p><p>Sin gastos registrados</p>
            </div>
          )}
          {gastosList.map(m => <MovementRow key={m.id} m={m} onDelete={deleteMovement} onEdit={openEdit} products={products} />)}
        </div>
      )}

      {/* ── VENTAS ── */}
      {tab === "ventas" && (
        <div style={{ padding: "28px 20px 0" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
            <div>
              <h1 style={{ margin: 0, fontSize: 24, fontWeight: 700 }}>Ventas</h1>
              <p style={{ margin: "3px 0 0", fontSize: 13, color: "#999" }}>Total: {fmtAbs(ventasList.reduce((a,m) => a + m.amount, 0))}</p>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => exportToXLSX(ventasList, "ventas-kdp.xlsx", filterMonth ? `Ventas · ${MONTHS_ES[parseInt(filterMonth.split("-")[1])-1]} ${filterMonth.split("-")[0]}` : `Ventas · Histórico completo ${new Date().getFullYear()}`)}
                style={{ background: "#f0f0f0", color: "#555", border: "none", borderRadius: 22,
                  padding: "10px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>📥 Excel</button>
              <button onClick={() => { setModal("venta"); setForm({ date: now.toISOString().slice(0,10), concept: activeProduct?.name }); }}
                style={{ background: "#1a7a4a", color: "#fff", border: "none", borderRadius: 22,
                  padding: "10px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>+ Nueva venta</button>
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, margin: "16px 0 14px", alignItems: "center" }}>
            <div style={{ flex: 1, background: "#fff", borderRadius: 22, padding: "6px 14px",
              fontSize: 13, color: "#888", boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
              {filterMonth ? `Filtrando: ${MONTHS_ES[parseInt(filterMonth.split("-")[1])-1]} ${filterMonth.split("-")[0]}` : "Todo El Histórico"}
            </div>
            <input type="month" value={filterMonth} onChange={e => setFilterMonth(e.target.value)}
              style={{ border: "1.5px solid #e8e8e8", borderRadius: 11, padding: "6px 12px", fontSize: 12,
                background: "#fff", outline: "none", fontFamily: "inherit", cursor: "pointer" }} />
            {filterMonth && <button onClick={() => setFilterMonth("")}
              style={{ background: "#f0f0f0", border: "none", borderRadius: 11, padding: "6px 10px", fontSize: 12, cursor: "pointer" }}>✕</button>}
          </div>

          {ventasList.length === 0 && (
            <div style={{ textAlign: "center", padding: "50px 0", color: "#ccc" }}>
              <p style={{ fontSize: 36 }}>💰</p><p>Sin ventas registradas</p>
            </div>
          )}
          {ventasList.map(m => <MovementRow key={m.id} m={m} onDelete={deleteMovement} onEdit={openEdit} products={products} />)}
        </div>
      )}

      {/* ── INFORMES ── */}
      {tab === "informes" && (
        <div style={{ padding: "28px 20px 0" }}>
          <h1 style={{ margin: "0 0 4px", fontSize: 24, fontWeight: 700 }}>Informes</h1>
          <p style={{ margin: "0 0 16px", fontSize: 12, color: "#aaa" }}>Consulta cualquier periodo o ejercicio</p>

          {/* Year selector */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
            <span style={{ fontSize: 13, color: "#555", fontWeight: 500 }}>Ejercicio fiscal</span>
            <select value={informeYear} onChange={e => setInformeYear(Number(e.target.value))}
              style={{ border: "1.5px solid #e8e8e8", borderRadius: 9, padding: "6px 10px", fontSize: 13,
                background: "#fff", outline: "none", fontFamily: "inherit" }}>
              {(availableYears.length ? availableYears : [String(now.getFullYear())]).map(y =>
                <option key={y} value={y}>{y}</option>)}
            </select>
          </div>

          {/* Beneficio del año */}
          <div style={{ background: yearStats.resultado >= 0 ? "#1a7a4a" : "#c0392b", borderRadius: 16,
            padding: "18px 20px", marginBottom: 14, boxShadow: "0 4px 14px rgba(0,0,0,0.12)" }}>
            <p style={{ margin: "0 0 4px", fontSize: 11, color: "rgba(255,255,255,0.7)", fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>
              📊 Beneficio del año {informeYear}
            </p>
            <p style={{ margin: 0, fontSize: 34, fontWeight: 700, color: "#fff", fontFamily: "'DM Mono', monospace", letterSpacing: -1 }}>
              {fmtSigned(yearStats.resultado)}
            </p>
            <p style={{ margin: "4px 0 0", fontSize: 12, color: "rgba(255,255,255,0.65)" }}>ROI {yearStats.roi.toFixed(0)}%</p>
          </div>

          {/* Ventas / Gastos año */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
            {[["📈","Ventas", yearStats.ingresos, "#1a7a4a"],["📉","Gastos", yearStats.gastos, "#c0392b"]].map(([ic,lb,v,c]) => (
              <div key={lb} style={{ background: "#fff", borderRadius: 13, padding: "13px 15px", boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
                <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 5 }}>
                  <span style={{ fontSize: 14 }}>{ic}</span>
                  <span style={{ fontSize: 12, color: "#aaa" }}>{lb}</span>
                </div>
                <p style={{ margin: 0, fontSize: 18, fontWeight: 700, color: c, fontFamily: "'DM Mono', monospace" }}>{fmtAbs(v)}</p>
              </div>
            ))}
          </div>

          {/* Comparativa mensual */}
          <h3 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px", color: "#1a1a1a" }}>% Comparativa mensual</h3>
          {monthsInYear.length === 0 && <p style={{ color: "#bbb", fontSize: 13 }}>Sin datos para {informeYear}</p>}
          {monthsInYear.map(({ ym, ingresos, gastos, resultado }) => (
            <div key={ym} onClick={() => setSelectedMonth(ym)} style={{ background: "#fff", borderRadius: 13, padding: "13px 15px", marginBottom: 8, boxShadow: "0 1px 4px rgba(0,0,0,0.05)", cursor: "pointer" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <span style={{ fontSize: 13, fontWeight: 600, color: "#1a1a1a" }}>{monthName(ym)}</span>
                <span style={{ fontSize: 13, fontWeight: 700, fontFamily: "'DM Mono', monospace",
                  color: resultado >= 0 ? "#1a7a4a" : "#c0392b" }}>{fmtSigned(resultado)}</span>
              </div>
              {/* Bar: ventas */}
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
                <span style={{ fontSize: 11, color: "#1a7a4a", width: 44, flexShrink: 0 }}>Ventas</span>
                <div style={{ flex: 1, background: "#f0f0f0", borderRadius: 4, height: 7 }}>
                  <div style={{ width: `${(ingresos/maxBar)*100}%`, background: "#1a7a4a", borderRadius: 4, height: 7, transition: "width 0.4s" }} />
                </div>
                <span style={{ fontSize: 11, color: "#1a7a4a", fontFamily: "'DM Mono', monospace", width: 70, textAlign: "right", flexShrink: 0 }}>{fmtAbs(ingresos)}</span>
              </div>
              {/* Bar: gastos */}
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 11, color: "#c0392b", width: 44, flexShrink: 0 }}>Gastos</span>
                <div style={{ flex: 1, background: "#f0f0f0", borderRadius: 4, height: 7 }}>
                  <div style={{ width: `${(gastos/maxBar)*100}%`, background: "#c0392b", borderRadius: 4, height: 7, transition: "width 0.4s" }} />
                </div>
                <span style={{ fontSize: 11, color: "#c0392b", fontFamily: "'DM Mono', monospace", width: 70, textAlign: "right", flexShrink: 0 }}>{fmtAbs(gastos)}</span>
              </div>
            </div>
          ))}

          {/* Rankings */}
          {/* Rankings: si hay 1 producto muestra resumen, si hay más muestra ranking */}
          <div style={{ marginTop: 18 }}>
            {products.length === 1 ? (
              <div>
                <h3 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px" }}>📊 Resumen del libro · {informeYear}</h3>
                {rankings.map(({ p, resultado, roi, ingresos, gastos }) => (
                  <div key={p.id} style={{ background: "#fff", borderRadius: 13, padding: "14px 15px", marginBottom: 7,
                    boxShadow: "0 1px 4px rgba(0,0,0,0.05)" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 11, marginBottom: 12 }}>
                      <div style={{ width: 34, height: 34, borderRadius: 9, background: p.color,
                        display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17 }}>{p.emoji}</div>
                      <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#1a1a1a",
                        whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", flex: 1 }}>{p.name}</p>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 7 }}>
                      {[["Ingresos", fmtAbs(ingresos), "#1a7a4a"],["Gastos", fmtAbs(gastos), "#c0392b"],
                        ["Resultado", fmtSigned(resultado), resultado>=0?"#1a7a4a":"#c0392b"],
                        ["ROI", roi.toFixed(0)+"%", roi>=0?"#1a7a4a":"#c0392b"]].map(([l,v,c]) => (
                        <div key={l} style={{ background: "#f7f5f0", borderRadius: 9, padding: "9px 8px" }}>
                          <p style={{ margin: 0, fontSize: 9, color: "#bbb", fontWeight: 700, textTransform: "uppercase" }}>{l}</p>
                          <p style={{ margin: "4px 0 0", fontSize: 12, fontWeight: 700, color: c, fontFamily: "'DM Mono', monospace" }}>{v}</p>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              (() => {
                const positivos = rankings.filter(r => r.resultado > 0);
                const negativos = [...rankings].reverse().filter(r => r.resultado < 0);
                return (
                  <>
                    <div style={{ marginBottom: 18 }}>
                      <h3 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px" }}>🏆 Más rentables · {informeYear}</h3>
                      {positivos.length === 0
                        ? <p style={{ fontSize: 13, color: "#bbb", fontStyle: "italic", padding: "10px 0" }}>Ningún libro en positivo aún</p>
                        : positivos.slice(0,3).map(({ p, resultado, roi }, i) => (
                          <div key={p.id} style={{ background: "#fff", borderRadius: 13, padding: "12px 15px", marginBottom: 7,
                            boxShadow: "0 1px 4px rgba(0,0,0,0.05)", display: "flex", alignItems: "center", gap: 11 }}>
                            <span style={{ fontSize: 16, fontWeight: 700, color: "#bbb", width: 20 }}>{i+1}</span>
                            <div style={{ width: 34, height: 34, borderRadius: 9, background: p.color,
                              display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17 }}>{p.emoji}</div>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#1a1a1a",
                                whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</p>
                              <p style={{ margin: 0, fontSize: 11, color: "#bbb" }}>ROI {roi.toFixed(0)}%</p>
                            </div>
                            <span style={{ fontSize: 14, fontWeight: 700, fontFamily: "'DM Mono', monospace", color: "#1a7a4a" }}>{fmtSigned(resultado)}</span>
                          </div>
                        ))
                      }
                    </div>
                    {negativos.length > 0 && (
                      <div style={{ marginBottom: 18 }}>
                        <h3 style={{ fontSize: 14, fontWeight: 700, margin: "0 0 10px" }}>📉 Menor rentabilidad · {informeYear}</h3>
                        {negativos.slice(0,3).map(({ p, resultado, roi }, i) => (
                          <div key={p.id} style={{ background: "#fff", borderRadius: 13, padding: "12px 15px", marginBottom: 7,
                            boxShadow: "0 1px 4px rgba(0,0,0,0.05)", display: "flex", alignItems: "center", gap: 11 }}>
                            <span style={{ fontSize: 16, fontWeight: 700, color: "#bbb", width: 20 }}>{i+1}</span>
                            <div style={{ width: 34, height: 34, borderRadius: 9, background: p.color,
                              display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17 }}>{p.emoji}</div>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#1a1a1a",
                                whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</p>
                              <p style={{ margin: 0, fontSize: 11, color: "#bbb" }}>ROI {roi.toFixed(0)}%</p>
                            </div>
                            <span style={{ fontSize: 14, fontWeight: 700, fontFamily: "'DM Mono', monospace", color: "#c0392b" }}>{fmtSigned(resultado)}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </>
                );
              })()
            )}
          </div>
          <div style={{ height: 20 }} />
        </div>
      )}

      {/* ── MODAL MES DETALLE ── */}
      {selectedMonth && (() => {
        const mvs = movements
          .filter(m => m.productId === activePid && m.date.startsWith(selectedMonth))
          .sort((a,b) => b.date.localeCompare(a.date));
        const s = calcStats(movements, activePid, selectedMonth);
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "flex-end", zIndex: 200 }}
            onClick={() => setSelectedMonth(null)}>
            <div style={{ background: "#f7f5f0", borderRadius: "22px 22px 0 0", padding: "24px 20px 44px",
              width: "100%", maxWidth: 420, margin: "0 auto", maxHeight: "80vh", overflowY: "auto" }}
              onClick={e => e.stopPropagation()}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700 }}>{monthName(selectedMonth)}</h3>
                <button onClick={() => setSelectedMonth(null)}
                  style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: "#aaa" }}>✕</button>
              </div>
              {/* Stats resumen */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 16 }}>
                {[["Ventas", fmtAbs(s.ingresos), "#1a7a4a"],["Gastos", fmtAbs(s.gastos), "#c0392b"],["ROI", s.roi.toFixed(0)+"%", s.roi>=0?"#1a7a4a":"#c0392b"]].map(([l,v,c]) => (
                  <div key={l} style={{ background: "#fff", borderRadius: 11, padding: "10px 12px", boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
                    <p style={{ margin: 0, fontSize: 10, color: "#bbb", fontWeight: 700, textTransform: "uppercase" }}>{l}</p>
                    <p style={{ margin: "4px 0 0", fontSize: 14, fontWeight: 700, color: c, fontFamily: "'DM Mono', monospace" }}>{v}</p>
                  </div>
                ))}
              </div>
              {/* Resultado */}
              <div style={{ background: s.resultado>=0?"#1a7a4a":"#c0392b", borderRadius: 13, padding: "12px 16px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 13, color: "rgba(255,255,255,0.8)" }}>Resultado</span>
                <span style={{ fontSize: 20, fontWeight: 700, color: "#fff", fontFamily: "'DM Mono', monospace" }}>{fmtSigned(s.resultado)}</span>
              </div>
              {/* Movimientos */}
              <h4 style={{ fontSize: 13, fontWeight: 700, color: "#888", textTransform: "uppercase", letterSpacing: 0.5, margin: "0 0 10px" }}>Movimientos</h4>
              {mvs.length === 0 && <p style={{ color: "#bbb", fontSize: 13 }}>Sin movimientos</p>}
              {mvs.map(m => (
                <div key={m.id} style={{ background: "#fff", borderRadius: 12, padding: "11px 14px", marginBottom: 7,
                  boxShadow: "0 1px 3px rgba(0,0,0,0.05)", display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ width: 34, height: 34, borderRadius: 9, flexShrink: 0,
                    background: m.type === "venta" ? "#e8f5e9" : "#fdecea",
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>
                    {m.type === "venta" ? "📈" : "📉"}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#1a1a1a",
                      whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{m.concept}</p>
                    <p style={{ margin: "2px 0 0", fontSize: 11, color: "#aaa" }}>{m.date}</p>
                  </div>
                  <span style={{ fontSize: 14, fontWeight: 700, fontFamily: "'DM Mono', monospace",
                    color: m.type === "venta" ? "#1a7a4a" : "#c0392b", flexShrink: 0 }}>
                    {m.type === "venta" ? "+" : "-"}{fmtAbs(m.amount)}
                  </span>
                </div>
              ))}
            </div>
          </div>
        );
      })()}

      {/* ── BOTTOM NAV ── */}}
      <div style={{ position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)", width: "100%",
        maxWidth: 420, background: "#fff", borderTop: "1px solid #f0f0f0", display: "flex",
        boxShadow: "0 -2px 12px rgba(0,0,0,0.06)", zIndex: 100 }}>
        {[["inicio","🏠","Inicio"],["productos","📦","Productos"],["gastos","📉","Gastos"],["ventas","📈","Ventas"],["informes","📊","Informes"]].map(([id,ic,lb]) => (
          <button key={id} onClick={() => setTab(id)}
            style={{ flex: 1, background: "none", border: "none", padding: "11px 4px 9px", cursor: "pointer",
              display: "flex", flexDirection: "column", alignItems: "center", gap: 3 }}>
            <span style={{ fontSize: 19 }}>{ic}</span>
            <span style={{ fontSize: 10, fontWeight: tab === id ? 700 : 400, color: tab === id ? "#1a1a1a" : "#ccc" }}>{lb}</span>
            {tab === id && <div style={{ width: 4, height: 4, borderRadius: 2, background: "#1a1a1a" }} />}
          </button>
        ))}
      </div>

      {/* ── MODAL VENTA ── */}
      {modal === "venta" && (() => {
        const unitPrice = parseFloat((form.amount || "0").replace(",", ".")) || 0;
        const units = parseInt(form.units || "1") || 1;
        const total = unitPrice * units;
        return (
        <Modal title="Nueva venta" onClose={() => { setModal(null); setForm({}); }}>
          <Field label="Fecha de cobro">
            <input type="date" value={form.date || ""} onChange={e => setF("date", e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Producto">
            <select value={form.concept || activeProduct?.name} onChange={e => setF("concept", e.target.value)} style={inputStyle}>
              {products.map(p => <option key={p.id} value={p.name}>{p.emoji} {p.name}</option>)}
            </select>
          </Field>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Field label="Importe por venta (€)">
              <input type="number" step="0.01" placeholder="0,00" value={form.amount || ""}
                onChange={e => setF("amount", e.target.value)} style={inputStyle} />
            </Field>
            <Field label="Nº de ventas">
              <input type="number" min="1" step="1" placeholder="1" value={form.units || ""}
                onChange={e => setF("units", e.target.value)} style={inputStyle} />
            </Field>
          </div>
          {/* Total calculado */}
          <div style={{ background: "#f5f0e8", borderRadius: 12, padding: "12px 16px", display: "flex",
            justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <span style={{ fontSize: 14, color: "#999" }}>Total</span>
            <span style={{ fontSize: 20, fontWeight: 700, color: "#1a7a4a", fontFamily: "'DM Mono', monospace" }}>
              {total.toFixed(2).replace(".", ",")} €
            </span>
          </div>
          {/* Mes de devengo */}
          <Field label="Mes al que pertenecen (regalías)">
            <select value={form.devengoMonth || ""} onChange={e => setF("devengoMonth", e.target.value)} style={inputStyle}>
              <option value="">— Mismo mes de cobro —</option>
              {MONTHS_ES.map((m, i) => {
                const val = `${String(i+1).padStart(2,"0")}/${now.getFullYear()}`;
                return <option key={i} value={val}>{m} {now.getFullYear()}</option>;
              })}
              {MONTHS_ES.map((m, i) => {
                const y = now.getFullYear() - 1;
                const val = `${String(i+1).padStart(2,"0")}/${y}`;
                return <option key={`${i}-prev`} value={val}>{m} {y}</option>;
              })}
            </select>
          </Field>
          <Field label="Notas (opcional)">
            <textarea placeholder="Ej. Venta hecha en abril, cobrada a finales de junio"
              value={form.notes || ""} onChange={e => setF("notes", e.target.value)}
              style={{ ...inputStyle, height: 72, resize: "none", lineHeight: 1.5 }} />
          </Field>
          <button onClick={() => {
            if (!total) return;
            const concept = form.concept || activeProduct?.name;
            setMovements(prev => [...prev, {
              id: Date.now(), productId: activePid, type: "venta", concept,
              amount: total,
              units: units,
              date: form.date || now.toISOString().slice(0,10),
              devengoMonth: form.devengoMonth || null,
              notes: form.notes || null,
            }]);
            setModal(null); setForm({});
          }} style={{ width: "100%", background: "#1a7a4a", color: "#fff", border: "none", borderRadius: 14,
            padding: "16px", fontSize: 15, fontWeight: 700, cursor: "pointer", marginTop: 4 }}>
            Guardar venta
          </button>
        </Modal>
        );
      })()}

      {/* ── MODAL GASTO ── */}
      {modal === "gasto" && (
        <Modal title="Nuevo gasto" onClose={() => { setModal(null); setForm({}); }}>
          <Field label="Categoría">
            <select value={form.concept || ""} onChange={e => setF("concept", e.target.value)} style={inputStyle}>
              <option value="">Seleccionar...</option>
              {CATS_GASTO.map(c => <option key={c} value={c}>{c}</option>)}
              <option value="__custom">Otro...</option>
            </select>
          </Field>
          {form.concept === "__custom" && (
            <Field label="Nombre del gasto">
              <input placeholder="Descripción" value={form.customConcept || ""}
                onChange={e => setF("customConcept", e.target.value)} style={inputStyle} />
            </Field>
          )}
          <Field label="Importe (€)">
            <input type="number" step="0.01" placeholder="0,00" value={form.amount || ""}
              onChange={e => setF("amount", e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Fecha">
            <input type="date" value={form.date || ""} onChange={e => setF("date", e.target.value)} style={inputStyle} />
          </Field>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 6 }}>
            <button onClick={() => { setModal(null); setForm({}); }}
              style={{ background: "#f0f0f0", border: "none", borderRadius: 12, padding: "14px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Cancelar</button>
            <button onClick={() => {
                const c = form.concept === "__custom" ? (form.customConcept || "Gasto") : (form.concept || "Gasto");
                setMovements(prev => [...prev, { id: Date.now(), productId: activePid, type: "gasto", concept: c, amount: parseFloat((form.amount||"0").replace(",",".")), date: form.date || now.toISOString().slice(0,10) }]);
                setModal(null); setForm({});
              }}
              style={{ background: "#c0392b", color: "#fff", border: "none", borderRadius: 12, padding: "14px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Guardar</button>
          </div>
        </Modal>
      )}

      {/* ── MODAL PRODUCTO ── */}
      {modal === "product" && (
        <Modal title="Nuevo producto" onClose={() => { setModal(null); setForm({}); }}>
          <Field label="Nombre del libro">
            <input placeholder="Título del libro" value={form.name || ""}
              onChange={e => setF("name", e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Emoji (icono)">
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 4 }}>
              {EMOJIS.map(e => (
                <button key={e} onClick={() => setF("emoji", e)}
                  style={{ fontSize: 22, background: form.emoji === e ? "#1a1a1a" : "#f0f0f0",
                    border: "none", borderRadius: 9, padding: "6px 10px", cursor: "pointer" }}>{e}</button>
              ))}
            </div>
          </Field>
          <Field label="Color de fondo">
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {COLORS.map(c => (
                <button key={c} onClick={() => setF("color", c)}
                  style={{ width: 30, height: 30, borderRadius: 8, background: c, border: form.color === c ? "2px solid #1a1a1a" : "2px solid transparent", cursor: "pointer" }} />
              ))}
            </div>
          </Field>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 8 }}>
            <button onClick={() => { setModal(null); setForm({}); }}
              style={{ background: "#f0f0f0", border: "none", borderRadius: 12, padding: "14px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Cancelar</button>
            <button onClick={saveProduct}
              style={{ background: "#1a1a1a", color: "#fff", border: "none", borderRadius: 12, padding: "14px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Crear</button>
          </div>
        </Modal>
      )}

      {/* ── MODAL EDITAR VENTA ── */}
      {modal === "edit-venta" && (
        <Modal title="Editar venta" onClose={() => { setModal(null); setForm({}); setEditingId(null); }}>
          <Field label="Concepto">
            <input value={form.concept || ""} onChange={e => setF("concept", e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Importe total (€)">
            <input type="number" step="0.01" value={form.amount || ""} onChange={e => setF("amount", e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Fecha de cobro">
            <input type="date" value={form.date || ""} onChange={e => setF("date", e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Mes al que pertenecen (regalías)">
            <select value={form.devengoMonth || ""} onChange={e => setF("devengoMonth", e.target.value)} style={inputStyle}>
              <option value="">— Mismo mes de cobro —</option>
              {MONTHS_ES.map((m, i) => {
                const val = `${String(i+1).padStart(2,"0")}/${now.getFullYear()}`;
                return <option key={i} value={val}>{m} {now.getFullYear()}</option>;
              })}
              {MONTHS_ES.map((m, i) => {
                const y = now.getFullYear() - 1;
                const val = `${String(i+1).padStart(2,"0")}/${y}`;
                return <option key={`${i}-prev`} value={val}>{m} {y}</option>;
              })}
            </select>
          </Field>
          <Field label="Notas (opcional)">
            <textarea value={form.notes || ""} onChange={e => setF("notes", e.target.value)}
              style={{ ...inputStyle, height: 72, resize: "none", lineHeight: 1.5 }} />
          </Field>
          <button onClick={updateMovement}
            style={{ width: "100%", background: "#1a7a4a", color: "#fff", border: "none", borderRadius: 14,
              padding: "16px", fontSize: 15, fontWeight: 700, cursor: "pointer", marginTop: 4 }}>
            Guardar cambios
          </button>
        </Modal>
      )}

      {/* ── MODAL EDITAR GASTO ── */}
      {modal === "edit-gasto" && (
        <Modal title="Editar gasto" onClose={() => { setModal(null); setForm({}); setEditingId(null); }}>
          <Field label="Concepto">
            <input value={form.concept || ""} onChange={e => setF("concept", e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Importe (€)">
            <input type="number" step="0.01" value={form.amount || ""} onChange={e => setF("amount", e.target.value)} style={inputStyle} />
          </Field>
          <Field label="Fecha">
            <input type="date" value={form.date || ""} onChange={e => setF("date", e.target.value)} style={inputStyle} />
          </Field>
          <button onClick={updateMovement}
            style={{ width: "100%", background: "#c0392b", color: "#fff", border: "none", borderRadius: 14,
              padding: "16px", fontSize: 15, fontWeight: 700, cursor: "pointer", marginTop: 4 }}>
            Guardar cambios
          </button>
        </Modal>
      )}
    </div>
  );
}
