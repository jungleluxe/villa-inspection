import { useState } from "react";
import * as XLSX from "xlsx";

const SECTIONS = [
  {
    id: "property", en: "Property Information", es: "Información de la Propiedad", icon: "🏡",
    fields: [
      { id: "villa_name", type: "text", en: "Villa Name", es: "Nombre de la Villa" },
      { id: "inspector_name", type: "text", en: "Inspector Name", es: "Nombre del Inspector" },
      { id: "inspection_date", type: "date", en: "Inspection Date", es: "Fecha de Inspección" },
      { id: "visit_type", type: "select", en: "Type of Visit", es: "Tipo de Visita", options: ["Initial Inspection / Inspección Inicial","Routine / Rutina","Pre-Owner Stay / Pre-Estancia Propietario","Follow-Up / Seguimiento"] },
    ]
  },
  {
    id: "water", en: "Water Systems", es: "Sistemas de Agua", icon: "💧",
    items: [
      { id: "pool_count", type: "count", en: "Number of Pools", es: "Número de Albercas" },
      { id: "pool_type", type: "select_item", en: "Pool Type", es: "Tipo de Alberca", options: ["Chlorine / Cloro","Saltwater / Sal","Both / Ambos"] },
      { id: "jacuzzi", type: "yesno", en: "Heated Jacuzzi", es: "Jacuzzi con Calefacción" },
      { id: "jacuzzi_working", type: "status", en: "Jacuzzi Working Condition", es: "Estado del Jacuzzi", dependsOn: "jacuzzi" },
      { id: "water_softener", type: "yesno", en: "Water Softener System", es: "Suavizador de Agua" },
      { id: "water_heater", type: "yesno", en: "Water Heater", es: "Calentador de Agua" },
      { id: "water_heater_type", type: "select_item", en: "Water Heater Type", es: "Tipo de Calentador", options: ["Solar / Solar","Gas / Gas","Electric / Eléctrico","Hybrid / Híbrido"] },
      { id: "sprinkler", type: "yesno", en: "Sprinkler / Irrigation System", es: "Sistema de Riego / Aspersores" },
    ]
  },
  {
    id: "gas", en: "Gas & Propane", es: "Gas y Propano", icon: "🔥",
    items: [
      { id: "propane_tank", type: "yesno", en: "Propane Tank on Property", es: "Tanque de Propano en la Propiedad" },
      { id: "propane_level", type: "status", en: "Propane Tank Level", es: "Nivel del Tanque de Propano", dependsOn: "propane_tank" },
      { id: "gas_stove", type: "yesno", en: "Gas Stove", es: "Estufa de Gas" },
      { id: "gas_bbq", type: "yesno", en: "Gas BBQ / Grill", es: "Asador / Parrilla de Gas" },
    ]
  },
  {
    id: "hvac", en: "AC & Ventilation", es: "Aire Acondicionado y Ventilación", icon: "❄️",
    items: [
      { id: "ac_count", type: "count", en: "Total Number of AC Units", es: "Número Total de Aires Acondicionados" },
      { id: "ac_brand", type: "text_item", en: "AC Brand(s)", es: "Marca(s) del Aire Acondicionado" },
      { id: "ac_condition", type: "status", en: "General AC Condition", es: "Estado General de los Aires" },
      { id: "ceiling_fans", type: "count", en: "Ceiling Fans", es: "Ventiladores de Techo" },
    ]
  },
  {
    id: "appliances", en: "Appliances & Equipment", es: "Electrodomésticos y Equipos", icon: "🍳",
    items: [
      { id: "washer", type: "yesno", en: "Washing Machine", es: "Lavadora" },
      { id: "dryer", type: "yesno", en: "Dryer", es: "Secadora" },
      { id: "dishwasher", type: "yesno", en: "Dishwasher", es: "Lavavajillas" },
      { id: "icemaker", type: "yesno", en: "Ice Maker (Standalone)", es: "Máquina de Hielo" },
      { id: "wine_fridge", type: "yesno", en: "Wine Refrigerator", es: "Vinoteca" },
      { id: "outdoor_kitchen", type: "yesno", en: "Outdoor Kitchen", es: "Cocina Exterior" },
      { id: "coffee_machine", type: "yesno", en: "Specialty Coffee Machine", es: "Máquina de Café Especial" },
      { id: "other_appliances", type: "text_item", en: "Other Specialty Appliances", es: "Otros Electrodomésticos Especiales" },
    ]
  },
  {
    id: "internet", en: "Internet & Technology", es: "Internet y Tecnología", icon: "📶",
    items: [
      { id: "internet_provider", type: "select_item", en: "Internet Provider", es: "Proveedor de Internet", options: ["ABIX","GIGNET","Telmex","Other / Otro"] },
      { id: "router_location", type: "text_item", en: "Router Location", es: "Ubicación del Router" },
      { id: "wifi_extenders", type: "yesno", en: "WiFi Extenders / Mesh System", es: "Extensores de WiFi / Sistema Mesh" },
      { id: "extender_count", type: "count", en: "Number of Extenders", es: "Número de Extensores", dependsOn: "wifi_extenders" },
      { id: "smart_locks", type: "yesno", en: "Smart Locks / Keypad Entry", es: "Cerraduras Inteligentes" },
      { id: "smart_home", type: "yesno", en: "Smart Home System", es: "Sistema Casa Inteligente" },
      { id: "security_cameras", type: "yesno", en: "Security Cameras", es: "Cámaras de Seguridad" },
    ]
  },
  {
    id: "outdoor", en: "Outdoor & Grounds", es: "Exteriores y Jardines", icon: "🌿",
    items: [
      { id: "garden_service", type: "yesno", en: "Garden Maintenance Required", es: "Mantenimiento de Jardín Requerido" },
      { id: "garden_frequency", type: "select_item", en: "Garden Service Frequency", es: "Frecuencia del Servicio de Jardín", options: ["Weekly / Semanal","Bi-weekly / Quincenal","Monthly / Mensual"], dependsOn: "garden_service" },
      { id: "tall_windows", type: "yesno", en: "Tall / High Windows", es: "Ventanas Altas" },
      { id: "window_cleaning_cycle", type: "select_item", en: "Window Cleaning Frequency", es: "Frecuencia de Limpieza de Ventanas", options: ["Every 1 month / Cada mes","Every 3 months / Cada 3 meses","Every 6 months / Cada 6 meses"], dependsOn: "tall_windows" },
      { id: "outdoor_shower", type: "yesno", en: "Outdoor Shower", es: "Regadera Exterior" },
      { id: "palapa", type: "yesno", en: "Palapa / Thatched Structure", es: "Palapa / Estructura de Palma" },
      { id: "terrace_furniture", type: "status", en: "Terrace / Outdoor Furniture Condition", es: "Estado de Muebles de Terraza" },
    ]
  },
  {
    id: "immediate", en: "Immediate Service Required", es: "Servicio Inmediato Requerido", icon: "🚨",
    note_en: "List everything that needs a vendor visit NOW. PHOTOS ARE REQUIRED for every item.",
    note_es: "Anota todo lo que necesita atención AHORA. SE REQUIEREN FOTOS de cada punto.",
    items: [
      { id: "imm_1", type: "immediate", en: "Issue #1", es: "Problema #1" },
      { id: "imm_2", type: "immediate", en: "Issue #2", es: "Problema #2" },
      { id: "imm_3", type: "immediate", en: "Issue #3", es: "Problema #3" },
      { id: "imm_4", type: "immediate", en: "Issue #4", es: "Problema #4" },
      { id: "imm_5", type: "immediate", en: "Issue #5", es: "Problema #5" },
    ]
  }
];

const STATUS_OPTIONS = [
  { value: "ok", label: "✅ OK", color: "#22c55e" },
  { value: "attention", label: "⚠️ Needs Attention / Necesita Atención", color: "#f59e0b" },
  { value: "urgent", label: "🔴 Urgent – Vendor Required / Urgente", color: "#ef4444" },
];

function exportToExcel(data) {
  const wb = XLSX.utils.book_new();
  const rows = [
    ["JUNGLE LUXE — VILLA OPERATIONS INSPECTION",""],
    [""],
    ["Villa Name / Nombre de la Villa", data["villa_name"] || ""],
    ["Inspector", data["inspector_name"] || ""],
    ["Date / Fecha", data["inspection_date"] || ""],
    ["Visit Type / Tipo de Visita", data["visit_type"] || ""],
    [""],
  ];
  SECTIONS.forEach(s => {
    if (s.id === "property") return;
    rows.push([`--- ${s.en.toUpperCase()} / ${s.es.toUpperCase()} ---`, ""]);
    (s.fields || s.items || []).forEach(item => {
      if (item.type === "immediate") {
        const desc = data[item.id + "_desc"];
        if (desc) {
          rows.push([`${item.en}`, desc]);
          rows.push(["  Priority / Prioridad", data[item.id + "_priority"] || ""]);
          rows.push(["  Vendor / Proveedor", data[item.id + "_vendor"] || ""]);
          rows.push(["  📸 Photo Required", "YES / SÍ"]);
        }
      } else {
        const val = data[item.id];
        if (val !== undefined && val !== "") {
          rows.push([`${item.en} / ${item.es}`, val]);
          if (data[item.id + "_notes"]) rows.push(["  Notes/Notas", data[item.id + "_notes"]]);
        }
      }
    });
    rows.push([""]);
  });
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 45 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, ws, "Inspection Summary");

  const urgentRows = [
    ["JUNGLE LUXE — IMMEDIATE SERVICE REQUIRED / SERVICIO INMEDIATO"],
    ["Villa", data["villa_name"] || ""], ["Date", data["inspection_date"] || ""],
    [""],
    ["#", "Issue / Problema", "Priority / Prioridad", "Vendor / Proveedor", "Photo Sent / Foto Enviada"],
  ];
  let c = 0;
  SECTIONS.find(s => s.id === "immediate").items.forEach(item => {
    const desc = data[item.id + "_desc"];
    if (desc) { c++; urgentRows.push([c, desc, data[item.id + "_priority"] || "", data[item.id + "_vendor"] || "", "☐"]); }
  });
  if (!c) urgentRows.push(["", "No immediate issues / Sin problemas inmediatos"]);
  const ws2 = XLSX.utils.aoa_to_sheet(urgentRows);
  ws2["!cols"] = [{ wch: 5 }, { wch: 45 }, { wch: 20 }, { wch: 25 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Immediate Issues");

  const filename = `JL_Inspection_${(data["villa_name"] || "Villa").replace(/\s+/g, "_")}_${data["inspection_date"] || new Date().toISOString().split("T")[0]}.xlsx`;
  XLSX.writeFile(wb, filename);
}

export default function App() {
  const [active, setActive] = useState(0);
  const [data, setData] = useState({});
  const [done, setDone] = useState(false);
  const set = (k, v) => setData(p => ({ ...p, [k]: v }));
  const get = (k, fb = "") => data[k] ?? fb;

  const progress = () => {
    let total = 0, filled = 0;
    SECTIONS.forEach(s => {
      (s.fields || []).forEach(f => { total++; if (get(f.id)) filled++; });
      (s.items || []).forEach(i => {
        if (i.type === "immediate") return;
        if (i.dependsOn && !get(i.dependsOn)) return;
        total++; if (get(i.id) !== "") filled++;
      });
    });
    return { total, filled, pct: total ? Math.round(filled / total * 100) : 0 };
  };

  const sectionPct = (s) => {
    let total = 0, filled = 0;
    (s.fields || []).forEach(f => { total++; if (get(f.id)) filled++; });
    (s.items || []).filter(i => i.type !== "immediate").forEach(i => {
      if (i.dependsOn && !get(i.dependsOn)) return;
      total++; if (get(i.id) !== "") filled++;
    });
    return total === 0 ? 100 : Math.round(filled / total * 100);
  };

  const inp = { width: "100%", background: "#1a1a1a", border: "1px solid #2e2e2e", borderRadius: 8, padding: "8px 12px", fontSize: 13, color: "#fff", outline: "none", boxSizing: "border-box" };
  const prog = progress();
  const sec = SECTIONS[active];

  const renderField = (f) => (
    <div key={f.id} style={{ marginBottom: 20 }}>
      <label style={{ display: "block", fontSize: 11, fontWeight: 700, letterSpacing: 2, textTransform: "uppercase", color: "#b9a06a", marginBottom: 2 }}>{f.en}</label>
      <label style={{ display: "block", fontSize: 10, color: "#666", fontStyle: "italic", marginBottom: 8 }}>{f.es}</label>
      {f.type === "text" && <input style={inp} value={get(f.id)} onChange={e => set(f.id, e.target.value)} placeholder={f.en} />}
      {f.type === "date" && <input type="date" style={inp} value={get(f.id)} onChange={e => set(f.id, e.target.value)} />}
      {f.type === "select" && <select style={inp} value={get(f.id)} onChange={e => set(f.id, e.target.value)}><option value="">— Select —</option>{f.options.map(o => <option key={o}>{o}</option>)}</select>}
    </div>
  );

  const renderItem = (item) => {
    if (item.dependsOn && !get(item.dependsOn)) return null;
    return (
      <div key={item.id} style={{ marginBottom: 20, paddingBottom: 20, borderBottom: "1px solid #1a1a1a" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
          <div style={{ flex: 1 }}>
            <p style={{ margin: "0 0 2px", fontWeight: 600, fontSize: 14, color: "#fff" }}>{item.en}</p>
            <p style={{ margin: 0, fontSize: 11, color: "#777", fontStyle: "italic" }}>{item.es}</p>
          </div>
          {item.type === "yesno" && (
            <div style={{ display: "flex", gap: 8 }}>
              {["yes", "no"].map(v => (
                <button key={v} onClick={() => set(item.id, get(item.id) === v ? "" : v)}
                  style={{ padding: "6px 14px", borderRadius: 8, border: `1px solid ${get(item.id) === v ? (v === "yes" ? "#22c55e" : "#ef4444") : "#2e2e2e"}`, background: get(item.id) === v ? (v === "yes" ? "#22c55e22" : "#ef444422") : "transparent", color: get(item.id) === v ? (v === "yes" ? "#22c55e" : "#ef4444") : "#666", fontWeight: 700, fontSize: 12, cursor: "pointer" }}>
                  {v === "yes" ? "Sí / Yes" : "No"}
                </button>
              ))}
            </div>
          )}
          {item.type === "count" && (
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <button onClick={() => set(item.id, Math.max(0, (get(item.id) || 0) - 1))} style={{ width: 32, height: 32, borderRadius: "50%", border: "1px solid #2e2e2e", background: "transparent", color: "#b9a06a", fontSize: 18, cursor: "pointer" }}>−</button>
              <span style={{ width: 32, textAlign: "center", fontWeight: 700, fontSize: 18, color: "#fff" }}>{get(item.id) || 0}</span>
              <button onClick={() => set(item.id, (get(item.id) || 0) + 1)} style={{ width: 32, height: 32, borderRadius: "50%", border: "1px solid #2e2e2e", background: "transparent", color: "#b9a06a", fontSize: 18, cursor: "pointer" }}>+</button>
            </div>
          )}
        </div>
        {item.type === "status" && (
          <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 6 }}>
            {STATUS_OPTIONS.map(s => (
              <button key={s.value} onClick={() => set(item.id, get(item.id) === s.value ? "" : s.value)}
                style={{ textAlign: "left", padding: "8px 12px", borderRadius: 8, border: `1px solid ${get(item.id) === s.value ? s.color : "#2e2e2e"}`, background: get(item.id) === s.value ? s.color + "22" : "transparent", color: get(item.id) === s.value ? s.color : "#888", fontSize: 12, cursor: "pointer" }}>
                {s.label}
              </button>
            ))}
          </div>
        )}
        {item.type === "select_item" && <div style={{ marginTop: 10 }}><select style={inp} value={get(item.id)} onChange={e => set(item.id, e.target.value)}><option value="">— Select —</option>{item.options.map(o => <option key={o}>{o}</option>)}</select></div>}
        {item.type === "text_item" && <div style={{ marginTop: 10 }}><input style={inp} value={get(item.id)} onChange={e => set(item.id, e.target.value)} placeholder="Notes / Notas..." /></div>}
        {item.type === "immediate" && (
          <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 8 }}>
            <input style={inp} value={get(item.id + "_desc")} onChange={e => set(item.id + "_desc", e.target.value)} placeholder="Describe the issue / Describe el problema" />
            <div style={{ display: "flex", gap: 8 }}>
              <div style={{ flex: 1, border: "1px dashed #ef4444", borderRadius: 8, padding: "8px 12px", fontSize: 12, color: "#ef4444", textAlign: "center" }}>📸 Photo Required / Foto Requerida</div>
              <select style={{ ...inp, flex: 1 }} value={get(item.id + "_priority")} onChange={e => set(item.id + "_priority", e.target.value)}>
                <option value="">Priority / Prioridad</option>
                <option>Same Day / Mismo Día</option>
                <option>48 hrs</option>
                <option>This Week / Esta Semana</option>
              </select>
            </div>
            <input style={inp} value={get(item.id + "_vendor")} onChange={e => set(item.id + "_vendor", e.target.value)} placeholder="Vendor needed / Proveedor requerido..." />
          </div>
        )}
        {["yesno", "status", "select_item"].includes(item.type) && (
          <textarea value={get(item.id + "_notes")} onChange={e => set(item.id + "_notes", e.target.value)} rows={1}
            placeholder="Notes / Notas (optional)"
            style={{ marginTop: 8, width: "100%", background: "#111", border: "1px solid #1a1a1a", borderRadius: 6, padding: "6px 10px", fontSize: 11, color: "#aaa", outline: "none", resize: "none", boxSizing: "border-box" }} />
        )}
      </div>
    );
  };

  if (done) {
    const urgent = SECTIONS.find(s => s.id === "immediate").items.filter(i => get(i.id + "_desc"));
    return (
      <div style={{ background: "#0d0d0d", minHeight: "100vh", fontFamily: "Georgia, serif", display: "flex", alignItems: "center", justifyContent: "center", padding: 32 }}>
        <div style={{ maxWidth: 480, width: "100%", textAlign: "center" }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>✅</div>
          <h2 style={{ color: "#b9a06a", fontSize: 22, margin: "0 0 4px" }}>Inspection Complete</h2>
          <p style={{ color: "#666", fontSize: 13, fontStyle: "italic", margin: "0 0 4px" }}>Inspección Completada</p>
          <p style={{ color: "#444", fontSize: 12, marginBottom: 32 }}>{get("villa_name") || "—"} · {get("inspector_name") || "—"} · {get("inspection_date") || "—"}</p>
          {urgent.length > 0 && (
            <div style={{ background: "#1a0000", border: "1px solid #ef4444", borderRadius: 12, padding: 20, marginBottom: 24, textAlign: "left" }}>
              <p style={{ color: "#ef4444", fontWeight: "bold", fontSize: 14, margin: "0 0 4px" }}>🚨 {urgent.length} Immediate Issue(s) — Send photos to Oscar NOW</p>
              <p style={{ color: "#f87171", fontSize: 11, fontStyle: "italic", margin: "0 0 12px" }}>Envía fotos a Oscar AHORA</p>
              {urgent.map((i, idx) => <div key={i.id} style={{ color: "#fca5a5", fontSize: 12, marginBottom: 6 }}>{idx + 1}. {get(i.id + "_desc")} — <span style={{ color: "#f87171" }}>{get(i.id + "_priority") || "TBD"}</span></div>)}
            </div>
          )}
          <div style={{ display: "flex", gap: 12, flexDirection: "column" }}>
            <button onClick={() => exportToExcel(data)} style={{ padding: 14, background: "#22c55e", border: "none", borderRadius: 10, color: "#fff", fontWeight: "bold", fontSize: 15, cursor: "pointer" }}>
              ⬇️ Download Excel Report / Descargar Reporte Excel
            </button>
            <button onClick={() => { setData({}); setDone(false); setActive(0); }} style={{ padding: 12, background: "transparent", border: "1px solid #2e2e2e", borderRadius: 10, color: "#888", fontSize: 13, cursor: "pointer" }}>
              New Inspection / Nueva Inspección
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ background: "#0d0d0d", minHeight: "100vh", fontFamily: "Georgia, serif", color: "#fff" }}>
      <div style={{ background: "#111", borderBottom: "1px solid #1e1e1e", padding: "14px 20px", position: "sticky", top: 0, zIndex: 10 }}>
        <div style={{ maxWidth: 640, margin: "0 auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div>
              <h1 style={{ margin: 0, fontSize: 14, fontWeight: "bold", color: "#b9a06a", letterSpacing: 3, textTransform: "uppercase" }}>Jungle Luxe</h1>
              <p style={{ margin: 0, fontSize: 10, color: "#555" }}>Villa Operations Inspection · Inspección Operativa</p>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: 18, fontWeight: "bold", color: prog.pct === 100 ? "#22c55e" : "#b9a06a" }}>{prog.pct}%</div>
              <div style={{ fontSize: 10, color: "#555" }}>{prog.filled}/{prog.total}</div>
            </div>
          </div>
          <div style={{ background: "#1a1a1a", borderRadius: 4, height: 3 }}>
            <div style={{ background: prog.pct === 100 ? "#22c55e" : "#b9a06a", height: 3, borderRadius: 4, width: `${prog.pct}%`, transition: "width 0.3s" }} />
          </div>
        </div>
      </div>
      <div style={{ background: "#111", borderBottom: "1px solid #1a1a1a", overflowX: "auto" }}>
        <div style={{ display: "flex", maxWidth: 640, margin: "0 auto", padding: "0 16px" }}>
          {SECTIONS.map((s, idx) => (
            <button key={s.id} onClick={() => setActive(idx)}
              style={{ background: "transparent", border: "none", borderBottom: `2px solid ${active === idx ? "#b9a06a" : "transparent"}`, padding: "10px 8px 8px", cursor: "pointer", fontSize: 18, position: "relative" }}>
              {s.icon}
              {sectionPct(s) === 100 && <span style={{ position: "absolute", top: 6, right: 2, fontSize: 8, color: "#22c55e" }}>✓</span>}
            </button>
          ))}
        </div>
      </div>
      <div style={{ maxWidth: 640, margin: "0 auto", padding: "24px 16px 80px" }}>
        <h2 style={{ fontSize: 18, fontWeight: "bold", color: "#fff", margin: "0 0 2px" }}>{sec.icon} {sec.en}</h2>
        <p style={{ fontSize: 12, color: "#666", fontStyle: "italic", margin: "0 0 24px" }}>{sec.es}</p>
        {sec.note_en && (
          <div style={{ background: "#1a0000", border: "1px solid #ef444466", borderRadius: 10, padding: "14px 18px", marginBottom: 24 }}>
            <p style={{ color: "#fca5a5", fontWeight: "bold", fontSize: 13, margin: "0 0 4px" }}>⚠️ {sec.note_en}</p>
            <p style={{ color: "#f87171", fontSize: 11, fontStyle: "italic", margin: 0 }}>{sec.note_es}</p>
          </div>
        )}
        {(sec.fields || []).map(renderField)}
        {(sec.items || []).map(renderItem)}
        <div style={{ display: "flex", gap: 10, marginTop: 28 }}>
          {active > 0 && <button onClick={() => setActive(a => a - 1)} style={{ flex: 1, padding: 12, background: "transparent", border: "1px solid #2e2e2e", borderRadius: 10, color: "#888", fontSize: 13, cursor: "pointer" }}>← Back / Atrás</button>}
          {active < SECTIONS.length - 1
            ? <button onClick={() => setActive(a => a + 1)} style={{ flex: 2, padding: 12, background: "#b9a06a", border: "none", borderRadius: 10, color: "#0d0d0d", fontWeight: "bold", fontSize: 14, cursor: "pointer" }}>Next / Siguiente →</button>
            : <button onClick={() => setDone(true)} style={{ flex: 2, padding: 12, background: "#22c55e", border: "none", borderRadius: 10, color: "#fff", fontWeight: "bold", fontSize: 14, cursor: "pointer" }}>✓ Submit / Enviar</button>
          }
        </div>
      </div>
    </div>
  );
}
