import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from "recharts";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { CONTRATO_TEMPLATE_B64 } from "./contratoTemplate.js";
import {
  Home, Users, Wallet, AlertTriangle, KeyRound, Plus, X, Pencil, Trash2,
  Upload, Download, ShieldCheck, ShieldAlert, DoorOpen, DoorClosed,
  ChevronLeft, ChevronRight, Loader2, Check, RefreshCw, WifiOff, History,
  LayoutGrid, RotateCcw, FileText, Paperclip, LogOut, FileSignature, Sparkles
} from "lucide-react";

/* ------------------------------------------------------------------ */
/* Utilidades                                                          */
/* ------------------------------------------------------------------ */

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const MESES_CORTOS = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
const GASTOS_FIJOS_KEYS = ["luz", "agua", "gas", "limpieza", "internet", "ibi", "comunidad"];
const GASTOS_FIJOS_LABELS = { luz: "Luz", agua: "Agua", gas: "Gas", limpieza: "Limpieza", internet: "Internet", ibi: "IBI", comunidad: "Comunidad de Propietarios" };
const REPARACIONES_CONCEPTOS = ["Electricidad", "Fontanería", "Albañilería", "Carpintería", "Pintura", "Electrodomésticos", "Muebles", "Ropa de Cama", "Cocina"];
const OTROS_CONCEPTOS = ["Notaría", "Registro de la Propiedad", "ITP", "Otros Impuestos"];
const ROOM_LABELS = ["A1", "A2", "B1", "B2", "C1", "C2", "D1", "E1", "E2", "E3", "F1", "F2"];
const TOTAL_HABITACIONES = ROOM_LABELS.length;

/* Escala combinada de referencia (estatal + autonómica media aproximada) IRPF 2025/2026.
   Los tipos autonómicos reales varían según la Comunidad Autónoma: esto es solo orientativo. */
const IRPF_TRAMOS = [
  { hasta: 12450, tipo: 0.19 },
  { hasta: 20200, tipo: 0.24 },
  { hasta: 35200, tipo: 0.30 },
  { hasta: 60000, tipo: 0.37 },
  { hasta: 300000, tipo: 0.45 },
  { hasta: Infinity, tipo: 0.47 }
];
const REDUCCIONES_ALQUILER = [
  { value: 0, label: "0% — Sin reducción" },
  { value: 0.5, label: "50% — Reducción general (vivienda habitual del inquilino)" },
  { value: 0.6, label: "60% — Rehabilitación reciente de la vivienda" },
  { value: 0.7, label: "70% — Inquilino de 18–35 años o entidad social, en zona tensionada" },
  { value: 0.9, label: "90% — Zona tensionada con rebaja de renta ≥5%" }
];

function calcIrpfProgresivo(base) {
  if (base <= 0) return 0;
  let cuota = 0;
  let anterior = 0;
  for (const tramo of IRPF_TRAMOS) {
    if (base > tramo.hasta) {
      cuota += (tramo.hasta - anterior) * tramo.tipo;
      anterior = tramo.hasta;
    } else {
      cuota += (base - anterior) * tramo.tipo;
      break;
    }
  }
  return cuota;
}

function pad2(n) { return String(n).padStart(2, "0"); }
function ymKey(y, m) { return `${y}-${pad2(m)}`; }
function todayDate() { return new Date(); }
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
function toDate(iso) { return iso ? new Date(iso + "T00:00:00") : null; }
function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("es-ES", { style: "currency", currency: "EUR", maximumFractionDigits: 0 });
}
function fmtDate(iso) {
  if (!iso) return "—";
  const d = toDate(iso);
  if (!d || isNaN(d.getTime())) return "—";
  return d.toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric" });
}
function effectiveEnd(t) {
  return (t.renovado && t.nuevaFechaFin) ? t.nuevaFechaFin : t.fechaFin;
}
function overlapsMonth(t, ym) {
  const [y, m] = ym.split("-").map(Number);
  const monthStart = new Date(y, m - 1, 1);
  const monthEnd = new Date(y, m, 0);
  const start = toDate(t.fechaInicio);
  const end = toDate(effectiveEnd(t));
  if (start && start > monthEnd) return false;
  if (end && end < monthStart) return false;
  return true;
}
function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }
function daysInMonth(y, m) { return new Date(y, m, 0).getDate(); } // m: 1-12

/* ------------------------------------------------------------------ */
/* Generación de contrato de alquiler (Word)                            */
/* ------------------------------------------------------------------ */

const MESES_FIRMA = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

function addDaysISO(iso, days) {
  const d = toDate(iso);
  if (!d) return "";
  d.setDate(d.getDate() + days);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function fechaFirmaPartes(iso) {
  const d = toDate(iso);
  if (!d) return { dia: "", mes: "", anio: "" };
  return { dia: String(d.getDate()), mes: MESES_FIRMA[d.getMonth()], anio: String(d.getFullYear()) };
}

/* Convierte un número entero (0-999.999) a su forma en letras, en mayúsculas, para uso legal/contractual */
const UNIDADES = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE"];
const DECENAS = ["", "", "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"];
const CENTENAS = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"];

function numeroALetras(n) {
  n = Math.floor(Math.abs(Number(n) || 0));
  if (n === 0) return "CERO";
  if (n === 100) return "CIEN";

  function tresDigitos(num) {
    if (num === 0) return "";
    if (num === 100) return "CIEN";
    let partes = [];
    const c = Math.floor(num / 100);
    const resto = num % 100;
    if (c > 0) partes.push(CENTENAS[c]);
    if (resto > 0) {
      if (resto <= 20) {
        partes.push(UNIDADES[resto]);
      } else {
        const d = Math.floor(resto / 10);
        const u = resto % 10;
        if (d === 2) {
          partes.push(u > 0 ? `VEINTI${UNIDADES[u].toLowerCase() === "uno" ? "ÚN" : UNIDADES[u]}` : "VEINTE");
        } else {
          partes.push(u > 0 ? `${DECENAS[d]} Y ${UNIDADES[u]}` : DECENAS[d]);
        }
      }
    }
    return partes.join(" ");
  }

  if (n < 1000) return tresDigitos(n);

  if (n < 1000000) {
    const miles = Math.floor(n / 1000);
    const resto = n % 1000;
    const milesTxt = miles === 1 ? "MIL" : `${tresDigitos(miles)} MIL`;
    return resto > 0 ? `${milesTxt} ${tresDigitos(resto)}` : milesTxt;
  }
  return String(n); // fuera de rango razonable para una renta/fianza
}

/* "1.234,50" -> letras + céntimos, ej: "MIL DOSCIENTOS TREINTA Y CUATRO EUROS CON CINCUENTA CÉNTIMOS" */
function importeALetras(valor) {
  const num = Number(valor) || 0;
  const enteros = Math.floor(num);
  const centimos = Math.round((num - enteros) * 100);
  let texto = `${numeroALetras(enteros)} EUROS`;
  if (centimos > 0) {
    texto += ` CON ${numeroALetras(centimos)} CÉNTIMOS`;
  }
  return texto;
}

function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function escaparXml(texto) {
  return String(texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

async function generarContratoDocx(tenant) {
  const faltantes = [];
  if (!tenant.nombre?.trim()) faltantes.push("nombre");
  if (!tenant.apellidos?.trim()) faltantes.push("apellidos");
  if (!tenant.numeroDocumento?.trim()) faltantes.push("documento de identidad");
  if (!tenant.nacionalidad?.trim()) faltantes.push("nacionalidad");
  if (!tenant.telefono?.trim()) faltantes.push("teléfono");
  if (!tenant.correo?.trim()) faltantes.push("correo electrónico");
  if (!tenant.habitacion?.trim()) faltantes.push("habitación");
  if (!tenant.fechaInicio) faltantes.push("fecha de inicio de contrato");
  if (!tenant.fechaFin) faltantes.push("fecha de fin de contrato");
  if (!tenant.renta) faltantes.push("renta mensual");
  if (!tenant.fianzaImporte) faltantes.push("importe de la fianza");
  if (faltantes.length) {
    throw new Error(`Faltan datos en la ficha para generar el contrato: ${faltantes.join(", ")}.`);
  }

  const fechaFirmaISO = addDaysISO(tenant.fechaInicio, -3);
  const { dia, mes, anio } = fechaFirmaPartes(fechaFirmaISO);
  const renta = Number(tenant.renta) || 0;
  const fianza = Number(tenant.fianzaImporte) || 0;
  const meses = renta > 0 ? (Math.round((fianza / renta) * 10) / 10).toString().replace(".", ",") : "";

  const valores = {
    "@@NOMBRE_COMPLETO@@": `${tenant.nombre.trim()} ${tenant.apellidos.trim()}`,
    "@@NACIONALIDAD@@": tenant.nacionalidad.trim(),
    "@@DOCUMENTO_NUM@@": tenant.numeroDocumento.trim(),
    "@@TELEFONO@@": tenant.telefono.trim(),
    "@@CORREO@@": tenant.correo.trim(),
    "@@HABITACION@@": tenant.habitacion.trim(),
    "@@FECHA_INICIO@@": fmtDate(tenant.fechaInicio),
    "@@FECHA_FIN@@": fmtDate(tenant.fechaFin),
    "@@RENTA_LETRAS@@": importeALetras(renta),
    "@@RENTA_CIFRA@@": renta.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    "@@FIANZA_LETRAS@@": importeALetras(fianza),
    "@@FIANZA_CIFRA@@": fianza.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    "@@FIANZA_MESES@@": meses,
    "@@LUGAR_FIRMA@@": "Alcantarilla",
    "@@FECHA_FIRMA_DIA@@": dia,
    "@@FECHA_FIRMA_MES@@": mes,
    "@@FECHA_FIRMA_ANIO@@": anio,
  };

  const zip = await JSZip.loadAsync(base64ToUint8Array(CONTRATO_TEMPLATE_B64));
  let xml = await zip.file("word/document.xml").async("text");
  for (const [token, valor] of Object.entries(valores)) {
    xml = xml.split(token).join(escaparXml(valor));
  }
  zip.file("word/document.xml", xml);

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });

  const nombreArchivo = `Contrato_Hab_${tenant.habitacion}_${tenant.apellidos}`.replace(/[^a-zA-Z0-9_]/g, "_") + ".docx";
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombreArchivo;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function matchesRoom(habitacion, roomLabel) {
  const h = String(habitacion || "").trim().toUpperCase();
  if (!h) return false;
  return h === String(roomLabel).trim().toUpperCase();
}

/* ------------------------------------------------------------------ */
/* MODO DEMO: los documentos se guardan solo en memoria del navegador   */
/* (URL.createObjectURL), no en ningún servidor. Al recargar la página  */
/* se pierden, igual que el resto de los datos de la demo.              */
/* ------------------------------------------------------------------ */
const ROW_ID = "demo";
const DEMO_DOCS_STORE = new Map(); // path -> { file, blobUrl, name, size, uploadedAt }

function sanitizeFileName(name) {
  return String(name || "archivo").replace(/[^a-zA-Z0-9._-]/g, "_");
}

async function uploadDoc(path, file) {
  const existing = DEMO_DOCS_STORE.get(path);
  if (existing?.blobUrl) URL.revokeObjectURL(existing.blobUrl);
  const blobUrl = URL.createObjectURL(file);
  const meta = { path, name: file.name, size: file.size, uploadedAt: new Date().toISOString(), blobUrl };
  DEMO_DOCS_STORE.set(path, meta);
  return { path, name: file.name, size: file.size, uploadedAt: meta.uploadedAt };
}

async function getDocSignedUrl(path) {
  const meta = DEMO_DOCS_STORE.get(path);
  if (!meta) throw new Error("Documento no encontrado (recuerda que en la demo los archivos no se guardan al recargar la página).");
  return meta.blobUrl;
}

async function deleteDoc(path) {
  const existing = DEMO_DOCS_STORE.get(path);
  if (existing?.blobUrl) URL.revokeObjectURL(existing.blobUrl);
  DEMO_DOCS_STORE.delete(path);
}

function emptyTenant() {
  return {
    id: uid(),
    habitacion: "",
    nombre: "",
    apellidos: "",
    tipoDocumento: "DNI",
    numeroDocumento: "",
    nacionalidad: "",
    telefono: "",
    correo: "",
    fechaInicio: todayISO(),
    fechaFin: "",
    renovado: false,
    nuevaFechaFin: "",
    empadronado: false,
    renta: 0,
    fianzaImporte: 0,
    fechaPagoFianza: "",
    fechaDevolucionFianza: "",
    observaciones: "",
    activo: true,
    pagos: {},
    documentosContrato: [],
    documentosIdentidad: []
  };
}

function parseExcelDate(v) {
  if (!v && v !== 0) return "";
  if (v instanceof Date && !isNaN(v.getTime())) {
    return `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}`;
  }
  const s = String(v).trim();
  if (!s) return "";
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    return `${y}-${pad2(Number(mo))}-${pad2(Number(d))}`;
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const [, y, mo, d] = m;
    return `${y}-${pad2(Number(mo))}-${pad2(Number(d))}`;
  }
  return "";
}

function truthy(v) {
  return /^(s[ií]|x|1|true|yes)$/i.test(String(v || "").trim());
}

/* ------------------------------------------------------------------ */
/* Estilos (token system)                                              */
/* ------------------------------------------------------------------ */

const GlobalStyles = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@500;600;700;800&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

    .rg-root {
      --bg: #0a0e1f;
      --surface: #121834;
      --surface-alt: #1a2142;
      --border: #262c4d;
      --text: #e7e8f5;
      --text-dim: #9498b8;
      --accent: #6366f1;
      --accent-dim: #232a5c;
      --ok: #22c55e;
      --ok-dim: #123322;
      --warn: #e0a93d;
      --warn-dim: #3a2f18;
      --danger: #ef4444;
      --danger-dim: #3a1620;
      --info: #2f6fed;
      --info-dim: #16294f;
      --radius: 12px;
      font-family: 'Inter', sans-serif;
      background: var(--bg);
      color: var(--text);
      min-height: 100vh;
    }
    .rg-root * { box-sizing: border-box; }
    .rg-display { font-family: 'Manrope', sans-serif; }
    .rg-mono { font-family: 'IBM Plex Mono', monospace; }

    .rg-demo-banner {
      position: sticky; top: 0; z-index: 100;
      display: flex; align-items: center; justify-content: center; gap: 8px;
      background: linear-gradient(90deg, #4f46e5, #2f6fed 55%, #14b8a6);
      color: #ffffff; font-size: 12.5px; font-weight: 600; text-align: center;
      padding: 9px 16px;
    }
    .rg-demo-banner strong { font-weight: 800; }

    .rg-shell { display: flex; min-height: 100vh; }
    .rg-sidebar {
      width: 220px; flex-shrink: 0; background: #0d1015;
      border-right: 1px solid var(--border);
      display: flex; flex-direction: column; padding: 20px 14px;
      gap: 4px;
    }
    .rg-brand {
      display: flex; align-items: center; gap: 10px; padding: 6px 10px 22px 10px;
      color: var(--accent);
    }
    .rg-brand-mark {
      width: 34px; height: 34px; border-radius: 8px; background: var(--accent-dim);
      display: flex; align-items: center; justify-content: center; color: var(--accent);
      flex-shrink: 0;
    }
    .rg-nav-item {
      display: flex; align-items: center; gap: 10px;
      padding: 10px 12px; border-radius: 8px;
      color: var(--text-dim); cursor: pointer; font-size: 13.5px; font-weight: 500;
      transition: background .15s ease, color .15s ease;
      border: none; background: transparent; width: 100%; text-align: left;
    }
    .rg-nav-item:hover { background: var(--surface-alt); color: var(--text); }
    .rg-nav-item.active { background: var(--accent-dim); color: var(--accent); }
    .rg-nav-badge {
      margin-left: auto; background: var(--danger); color: #fff; font-size: 10.5px;
      font-weight: 700; border-radius: 999px; padding: 1px 7px; font-family: 'IBM Plex Mono', monospace;
    }
    .rg-sidebar-footer { margin-top: auto; padding: 10px; font-size: 11px; color: var(--text-dim); }

    .rg-main { flex: 1; padding: 28px 34px; overflow-x: hidden; max-width: 1280px; }

    .rg-topbar { display: flex; align-items: center; justify-content: space-between; margin-bottom: 22px; flex-wrap: wrap; gap: 14px; }
    .rg-h1 { font-family: 'Manrope', sans-serif; font-size: 26px; font-weight: 600; margin: 0; }
    .rg-sub { color: var(--text-dim); font-size: 13px; margin-top: 3px; }

    .rg-month-picker { display: flex; align-items: center; gap: 6px; background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 4px; }
    .rg-month-picker span { font-family: 'IBM Plex Mono', monospace; font-size: 13px; min-width: 118px; text-align: center; text-transform: capitalize; }
    .rg-icon-btn {
      background: transparent; border: none; color: var(--text-dim); cursor: pointer;
      width: 28px; height: 28px; display: flex; align-items: center; justify-content: center;
      border-radius: 6px; transition: background .15s ease, color .15s ease;
    }
    .rg-icon-btn:hover { background: var(--surface-alt); color: var(--accent); }

    .rg-card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); }
    .rg-grid-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 20px; }
    .rg-grid-cards.cols-3 { grid-template-columns: repeat(3, 1fr); }
    .rg-grid-cards.cols-5 { grid-template-columns: repeat(5, 1fr); }
    @media (max-width: 980px) { .rg-grid-cards, .rg-grid-cards.cols-3, .rg-grid-cards.cols-5 { grid-template-columns: repeat(2, 1fr); } }
    @media (max-width: 600px) { .rg-grid-cards, .rg-grid-cards.cols-3, .rg-grid-cards.cols-5 { grid-template-columns: 1fr; } }

    .rg-rooms-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); gap: 14px; }
    .rg-room-card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 16px; display: flex; flex-direction: column; }
    .rg-room-card.rg-room-free { border-style: dashed; }
    .rg-room-card.rg-room-occupied { border-color: var(--ok); }
    .rg-room-number { font-family: 'Manrope', sans-serif; font-size: 17px; font-weight: 600; margin-bottom: 8px; }

    .rg-stat { padding: 16px 18px; }
    .rg-stat-label { font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--text-dim); margin-bottom: 8px; }
    .rg-stat-value { font-family: 'IBM Plex Mono', monospace; font-size: 22px; font-weight: 600; }

    .rg-chart-card { padding: 18px 20px 8px 4px; margin-bottom: 18px; }
    .rg-chart-title { font-family: 'Manrope', sans-serif; font-size: 16px; font-weight: 600; padding: 0 16px; margin-bottom: 6px; }

    .rg-btn {
      font-family: 'Inter', sans-serif; background: var(--accent); color: #ffffff; border: none;
      border-radius: 8px; padding: 9px 15px; font-weight: 600; font-size: 13px; cursor: pointer;
      display: inline-flex; align-items: center; gap: 6px; transition: filter .15s ease; white-space: nowrap;
    }
    .rg-btn:hover { filter: brightness(1.12); }
    .rg-btn:focus-visible, .rg-icon-btn:focus-visible, .rg-nav-item:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
    .rg-btn-ghost { background: transparent; color: var(--text); border: 1px solid var(--border); }
    .rg-btn-ghost:hover { background: var(--surface-alt); filter: none; }
    .rg-btn-danger { background: var(--danger); color: #fff; }
    .rg-btn:disabled { opacity: .5; cursor: not-allowed; }

    .rg-input, .rg-select, .rg-textarea {
      background: var(--bg); border: 1px solid var(--border); color: var(--text);
      border-radius: 7px; padding: 8px 10px; font-size: 13.5px; font-family: 'Inter', sans-serif; width: 100%;
    }
    .rg-textarea { resize: vertical; min-height: 60px; font-family: 'Inter', sans-serif; }
    .rg-input:focus, .rg-select:focus, .rg-textarea:focus { outline: 2px solid var(--accent); outline-offset: 1px; }
    .rg-label { font-size: 11px; text-transform: uppercase; letter-spacing: .05em; color: var(--text-dim); margin-bottom: 5px; display: block; }
    .rg-field { margin-bottom: 14px; }
    .rg-check { display: flex; align-items: center; gap: 8px; font-size: 13px; cursor: pointer; }
    .rg-check input { width: 16px; height: 16px; accent-color: var(--accent); }

    .rg-badge { display: inline-flex; align-items: center; gap: 4px; padding: 3px 10px; border-radius: 999px; font-size: 11px; font-weight: 600; font-family: 'IBM Plex Mono', monospace; white-space: nowrap; }
    .rg-badge-ok { background: var(--ok-dim); color: var(--ok); }
    .rg-badge-danger { background: var(--danger-dim); color: var(--danger); }
    .rg-badge-warn { background: var(--warn-dim); color: var(--warn); }
    .rg-badge-info { background: var(--info-dim); color: var(--info); }
    .rg-badge-neutral { background: var(--surface-alt); color: var(--text-dim); }

    .rg-stamp {
      border: 1.5px solid currentColor; border-radius: 6px; transform: rotate(-2deg);
      display: inline-block; padding: 2px 9px; font-family: 'IBM Plex Mono', monospace;
      font-weight: 700; font-size: 10px; letter-spacing: .08em; text-transform: uppercase; cursor: pointer;
      background: transparent;
    }

    .rg-table-wrap { overflow-x: auto; }
    .rg-table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
    .rg-table th { text-align: left; color: var(--text-dim); font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; padding: 10px 12px; border-bottom: 1px solid var(--border); white-space: nowrap; }
    .rg-table td { padding: 11px 12px; border-bottom: 1px solid var(--border); vertical-align: middle; white-space: nowrap; }
    .rg-table tr:hover td { background: var(--surface-alt); }
    .rg-table tr:last-child td { border-bottom: none; }
    .rg-table th:last-child, .rg-table td:last-child {
      position: sticky; right: 0; background: var(--surface);
      box-shadow: -8px 0 10px -8px rgba(0,0,0,.35);
    }
    .rg-table tr:hover td:last-child { background: var(--surface-alt); }
    .rg-table tr:last-child td:last-child { border-bottom: none; }

    .rg-empty { padding: 40px 20px; text-align: center; color: var(--text-dim); }
    .rg-empty svg { margin-bottom: 10px; opacity: .5; }

    .rg-modal-overlay { position: fixed; inset: 0; background: rgba(8,10,13,.72); display: flex; align-items: flex-start; justify-content: center; z-index: 60; padding: 40px 16px; overflow-y: auto; }
    .rg-modal { background: var(--surface); border: 1px solid var(--border); border-radius: 14px; max-width: 760px; width: 100%; padding: 26px; margin-bottom: 40px; }
    .rg-modal-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 18px; }
    .rg-login-card { width: 100%; max-width: 360px; padding: 28px 26px; }
    .rg-modal-title { font-family: 'Manrope', sans-serif; font-size: 20px; font-weight: 600; }
    .rg-section-title { font-size: 12px; text-transform: uppercase; letter-spacing: .06em; color: var(--accent); font-weight: 700; margin: 20px 0 10px 0; padding-top: 14px; border-top: 1px solid var(--border); }
    .rg-section-title:first-of-type { border-top: none; padding-top: 0; margin-top: 0; }
    .rg-form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 14px; }
    @media (max-width: 620px) { .rg-form-grid { grid-template-columns: 1fr; } }

    .rg-pago-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 6px; }
    @media (max-width: 620px) { .rg-pago-grid { grid-template-columns: repeat(3, 1fr); } }
    .rg-pago-cell {
      border: 1px solid var(--border); border-radius: 7px; padding: 6px 4px; text-align: center;
      cursor: pointer; font-size: 11px; user-select: none; transition: background .12s ease;
    }
    .rg-pago-cell.paid { background: var(--ok-dim); border-color: var(--ok); color: var(--ok); }
    .rg-pago-cell.unpaid { background: var(--danger-dim); border-color: var(--danger); color: var(--danger); }
    .rg-pago-cell:hover { filter: brightness(1.15); }

    .rg-doc-block { background: var(--bg); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px; margin-bottom: 12px; }
    .rg-doc-block-header { display: flex; align-items: center; justify-content: space-between; gap: 10px; font-size: 12.5px; font-weight: 600; margin-bottom: 8px; }
    .rg-doc-empty { font-size: 12px; color: var(--text-dim); }
    .rg-doc-list { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 6px; }
    .rg-doc-list li { display: flex; align-items: center; gap: 8px; font-size: 12.5px; color: var(--text-dim); }
    .rg-doc-name { color: var(--text); cursor: pointer; text-decoration: underline; text-underline-offset: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 260px; }
    .rg-doc-name:hover { color: var(--accent); }
    .rg-doc-date { margin-left: auto; font-size: 11px; color: var(--text-dim); font-family: 'IBM Plex Mono', monospace; }

    .rg-factura-row { display: flex; align-items: center; gap: 6px; margin-top: 6px; }

    .rg-items-block { background: var(--bg); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; margin-bottom: 12px; }
    .rg-items-block-header { display: flex; align-items: center; justify-content: space-between; font-size: 13px; font-weight: 600; margin-bottom: 10px; }
    .rg-items-empty { font-size: 12.5px; color: var(--text-dim); margin-bottom: 8px; }
    .rg-item-row { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
    .rg-item-row .rg-select { flex: 1 1 auto; }
    .rg-item-amount { width: 110px; flex: 0 0 110px; }
    .rg-items-actions { display: flex; align-items: center; gap: 16px; flex-wrap: wrap; margin-top: 4px; }
    .rg-chip { display: inline-flex; align-items: center; gap: 5px; background: var(--accent-dim); color: var(--accent); border: none; border-radius: 999px; padding: 4px 10px; font-size: 11px; font-weight: 600; cursor: pointer; }
    .rg-btn-link { display: inline-flex; align-items: center; gap: 5px; background: transparent; border: none; color: var(--text-dim); font-size: 11px; cursor: pointer; margin-top: 6px; padding: 0; }
    .rg-btn-link:hover { color: var(--accent); }

    .rg-toast {
      position: fixed; bottom: 22px; left: 50%; transform: translateX(-50%);
      background: var(--surface); border: 1px solid var(--accent); color: var(--text);
      padding: 11px 20px; border-radius: 10px; font-size: 13px; z-index: 100;
      display: flex; align-items: center; gap: 8px; box-shadow: 0 8px 24px rgba(0,0,0,.4);
    }

    .rg-save-indicator { display: flex; align-items: center; gap: 6px; font-size: 11.5px; color: var(--text-dim); }
    .rg-save-indicator svg { animation: rg-spin 1s linear infinite; }
    @keyframes rg-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }

    .rg-alert-group { margin-bottom: 20px; }
    .rg-alert-group-header { display: flex; align-items: center; gap: 8px; font-family: 'Manrope', sans-serif; font-size: 16px; font-weight: 600; margin-bottom: 10px; }
    .rg-alert-item {
      display: flex; align-items: center; justify-content: space-between; gap: 12px;
      padding: 12px 16px; border-radius: 10px; margin-bottom: 8px; cursor: pointer;
      border: 1px solid var(--border); transition: background .15s ease;
    }
    .rg-alert-item:hover { background: var(--surface-alt); }
    .rg-alert-item-name { font-weight: 600; font-size: 13.5px; }
    .rg-alert-item-sub { font-size: 12px; color: var(--text-dim); }

    .rg-confirm-box { max-width: 400px; }

    @media (max-width: 800px) {
      .rg-shell { flex-direction: column; }
      .rg-sidebar { width: 100%; flex-direction: row; align-items: center; overflow-x: auto; padding: 12px; }
      .rg-brand { padding: 0 10px 0 0; }
      .rg-sidebar-footer { display: none; }
      .rg-main { padding: 20px 20px 90px 20px; }
    }

    .rg-mobile-logout { display: none; }
    @media (max-width: 800px) {
      .rg-mobile-logout {
        display: flex; align-items: center; justify-content: center; gap: 7px;
        position: fixed; left: 16px; right: 16px; bottom: 16px; z-index: 60;
        background: var(--surface); border: 1px solid var(--border); color: var(--danger);
        padding: 13px 16px; border-radius: 12px; font-size: 13.5px; font-weight: 600;
        box-shadow: 0 10px 28px rgba(0,0,0,.45); cursor: pointer;
      }
      .rg-mobile-logout:active { background: var(--surface-alt); }
    }
    @media (prefers-reduced-motion: reduce) {
      .rg-root * { transition: none !important; animation: none !important; }
    }
  `}</style>
);

/* ------------------------------------------------------------------ */
/* App                                                                  */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/* Pantalla de acceso (email + contraseña)                              */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/* Datos de ejemplo (MODO DEMO)                                        */
/* ------------------------------------------------------------------ */

const DEMO_MONTH = "2026-08";

function mesesPagados(desdeYm, hastaYm) {
  const pagos = {};
  const [y0, m0] = desdeYm.split("-").map(Number);
  const [y1, m1] = hastaYm.split("-").map(Number);
  let y = y0, m = m0;
  while (y < y1 || (y === y1 && m <= m1)) {
    pagos[ymKey(y, m)] = true;
    m++;
    if (m > 12) { m = 1; y++; }
  }
  return pagos;
}

function buildDemoTenants() {
  const base = [
    { habitacion: "A1", nombre: "Laura", apellidos: "Martín Ruiz", tipoDocumento: "DNI", numeroDocumento: "45678912A", nacionalidad: "Española", telefono: "612 345 678", correo: "laura.martin@example.com", empadronado: true, renta: 420, inicio: "2026-01-15", fin: "2027-02-28", pagoAgosto: true },
    { habitacion: "A2", nombre: "Carlos", apellidos: "Fernández Gil", tipoDocumento: "DNI", numeroDocumento: "52134987B", nacionalidad: "Española", telefono: "699 123 456", correo: "carlos.fernandez@example.com", empadronado: true, renta: 395, inicio: "2026-02-01", fin: "2026-10-31", pagoAgosto: false },
    { habitacion: "B1", nombre: "Ana", apellidos: "Sánchez López", tipoDocumento: "NIE", numeroDocumento: "Y1234567X", nacionalidad: "Colombiana", telefono: "633 987 654", correo: "ana.sanchez@example.com", empadronado: false, renta: 450, inicio: "2026-03-14", fin: "2027-03-14", pagoAgosto: true },
    { habitacion: "C1", nombre: "Sofía", apellidos: "Ibáñez Prat", tipoDocumento: "DNI", numeroDocumento: "48765123C", nacionalidad: "Española", telefono: "677 234 567", correo: "sofia.ibanez@example.com", empadronado: true, renta: 430, inicio: "2026-01-20", fin: "2026-09-20", pagoAgosto: true },
    { habitacion: "C2", nombre: "Miguel", apellidos: "Torres Vega", tipoDocumento: "DNI", numeroDocumento: "51987654D", nacionalidad: "Española", telefono: "622 345 123", correo: "miguel.torres@example.com", empadronado: true, renta: 410, inicio: "2026-02-25", fin: "2026-08-25", pagoAgosto: true },
    { habitacion: "E1", nombre: "Jon", apellidos: "Andueza Etxarri", tipoDocumento: "DNI", numeroDocumento: "30456789E", nacionalidad: "Española", telefono: "688 456 789", correo: "jon.andueza@example.com", empadronado: true, renta: 400, inicio: "2026-01-12", fin: "2026-12-12", pagoAgosto: true },
    { habitacion: "E2", nombre: "Elena", apellidos: "Ruiz Campos", tipoDocumento: "DNI", numeroDocumento: "47123890F", nacionalidad: "Española", telefono: "655 678 901", correo: "elena.ruiz@example.com", empadronado: false, renta: 415, inicio: "2026-04-01", fin: "2027-04-01", pagoAgosto: false },
    { habitacion: "E3", nombre: "Pablo", apellidos: "Núñez Soto", tipoDocumento: "NIE", numeroDocumento: "Z7654321Y", nacionalidad: "Venezolana", telefono: "611 789 012", correo: "pablo.nunez@example.com", empadronado: false, renta: 405, inicio: "2026-01-05", fin: "2027-01-05", pagoAgosto: true },
    { habitacion: "F2", nombre: "Marta", apellidos: "Gómez Díaz", tipoDocumento: "DNI", numeroDocumento: "49123456G", nacionalidad: "Española", telefono: "699 234 890", correo: "marta.gomez@example.com", empadronado: true, renta: 440, inicio: "2026-06-18", fin: "2027-06-18", pagoAgosto: true },
  ];

  const tenants = base.map(b => {
    const pagos = mesesPagados("2026-01", "2026-08");
    if (!b.pagoAgosto) delete pagos["2026-08"];
    return {
      id: uid(),
      habitacion: b.habitacion,
      nombre: b.nombre,
      apellidos: b.apellidos,
      tipoDocumento: b.tipoDocumento,
      numeroDocumento: b.numeroDocumento,
      nacionalidad: b.nacionalidad,
      telefono: b.telefono,
      correo: b.correo,
      empadronado: b.empadronado,
      activo: true,
      fechaInicio: b.inicio,
      fechaFin: b.fin,
      renovado: false,
      nuevaFechaFin: "",
      renta: b.renta,
      fianzaImporte: b.renta,
      fianzaTouched: false,
      fechaPagoFianza: b.inicio,
      fechaDevolucionFianza: "",
      pagos,
      documentosContrato: [],
      documentosIdentidad: [],
      observaciones: "",
    };
  });

  // Un inquilino ya dado de baja, para mostrar el Histórico con datos reales
  tenants.push({
    id: uid(),
    habitacion: "B2",
    nombre: "Diego",
    apellidos: "Castillo Moya",
    tipoDocumento: "DNI",
    numeroDocumento: "44567123H",
    nacionalidad: "Española",
    telefono: "666 111 222",
    correo: "diego.castillo@example.com",
    empadronado: true,
    activo: false,
    fechaInicio: "2025-06-01",
    fechaFin: "2026-06-01",
    renovado: false,
    nuevaFechaFin: "",
    renta: 425,
    fianzaImporte: 425,
    fianzaTouched: false,
    fechaPagoFianza: "2025-06-01",
    fechaDevolucionFianza: "2026-06-10",
    pagos: mesesPagados("2025-06", "2026-05"),
    documentosContrato: [],
    documentosIdentidad: [],
    observaciones: "Se marchó por traslado laboral. Fianza devuelta sin incidencias.",
  });

  return tenants;
}

function buildDemoExpenses() {
  const meses = ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06", "2026-07", "2026-08"];
  const expenses = {};
  meses.forEach((ym, idx) => {
    expenses[ym] = {
      luz: 130 + idx * 3,
      agua: 55 + idx,
      gas: 40,
      limpieza: 60,
      internet: 45,
      ibi: idx === 2 ? 310 : 0,
      comunidad: 95,
      reparacionesItems: [],
      otrosItems: [],
      facturas: {},
    };
  });
  // Un par de meses con reparaciones y otros gastos de ejemplo, para mostrar los desplegables
  expenses["2026-05"].reparacionesItems = [
    { id: uid(), concepto: "Fontanería", importe: 180 },
    { id: uid(), concepto: "Electricidad", importe: 75 },
  ];
  expenses["2026-08"].reparacionesItems = [
    { id: uid(), concepto: "Pintura", importe: 120 },
  ];
  expenses["2026-03"].otrosItems = [
    { id: uid(), concepto: "Otros Impuestos", importe: 45 },
  ];
  expenses._settings = { modo: "porcentaje", porcentaje: 15, fijo: 0 };
  return expenses;
}

/* ------------------------------------------------------------------ */
/* App                                                                  */
/* ------------------------------------------------------------------ */

export default function App() {
  const now = todayDate();
  const [tenants, setTenants] = useState(() => buildDemoTenants());
  const [expenses, setExpenses] = useState(() => buildDemoExpenses());
  const [view, setView] = useState("dashboard");
  const [loaded] = useState(true);
  const [saving] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState(DEMO_MONTH);
  const [showForm, setShowForm] = useState(false);
  const [editingTenant, setEditingTenant] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [confirmFree, setConfirmFree] = useState(null);
  const [toast, setToast] = useState(null);
  const [irpfReduccion, setIrpfReduccion] = useState(0.5);
  const fileInputRef = useRef(null);
  const toastTimer = useRef(null);

  const [selYear, selMonthNum] = selectedMonth.split("-").map(Number);

  function notify(msg) {
    setToast(msg);
    if (toastTimer.current) clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(null), 3800);
  }

  function reiniciarDemo() {
    setTenants(buildDemoTenants());
    setExpenses(buildDemoExpenses());
    setIrpfReduccion(0.5);
    setSelectedMonth(DEMO_MONTH);
    setView("dashboard");
    notify("Demo reiniciada con los datos de ejemplo originales.");
  }

  /* --------------------------- Cálculos --------------------------- */

  function incomeForMonth(ym) {
    return tenants.reduce((sum, t) => sum + (t.pagos && t.pagos[ym] ? (Number(t.renta) || 0) : 0), 0);
  }
  function occupancyForMonth(ym) {
    return tenants.filter(t => overlapsMonth(t, ym)).length;
  }
  function occupancyPctForMonth(ym) {
    return (occupancyForMonth(ym) / TOTAL_HABITACIONES) * 100;
  }
  function incomeEmpadronadosForMonth(ym) {
    return tenants.reduce((sum, t) => sum + (t.empadronado && t.pagos && t.pagos[ym] ? (Number(t.renta) || 0) : 0), 0);
  }
  function incomeEmpadronadosForYear(year) {
    let total = 0;
    for (let m = 1; m <= 12; m++) total += incomeEmpadronadosForMonth(ymKey(year, m));
    return total;
  }
  function fixedExpensesForMonth(ym) {
    const e = expenses[ym] || {};
    const simple = GASTOS_FIJOS_KEYS.reduce((s, k) => s + (Number(e[k]) || 0), 0);
    const reparaciones = (e.reparacionesItems || []).reduce((s, it) => s + (Number(it.importe) || 0), 0);
    const otros = (e.otrosItems || []).reduce((s, it) => s + (Number(it.importe) || 0), 0);
    return simple + reparaciones + otros;
  }
  function getGestionConfig() {
    return expenses._settings || { modo: "porcentaje", porcentaje: 15, fijo: 0 };
  }
  function managementFeeForMonth(ym) {
    const cfg = getGestionConfig();
    if (cfg.modo === "fijo") return Number(cfg.fijo) || 0;
    return incomeForMonth(ym) * ((Number(cfg.porcentaje) || 0) / 100);
  }
  function updateGestionConfig(patch) {
    setExpenses(prev => ({
      ...prev,
      _settings: { ...(prev._settings || { modo: "porcentaje", porcentaje: 15, fijo: 0 }), ...patch }
    }));
  }
  function totalExpensesForMonth(ym) {
    return fixedExpensesForMonth(ym) + managementFeeForMonth(ym);
  }
  function netProfitForMonth(ym) {
    return incomeForMonth(ym) - totalExpensesForMonth(ym);
  }
  function profitabilityForMonth(ym) {
    const inc = incomeForMonth(ym);
    if (inc <= 0) return 0;
    return (netProfitForMonth(ym) / inc) * 100;
  }

  const yearData = useMemo(() => {
    return MESES_CORTOS.map((label, idx) => {
      const ym = ymKey(selYear, idx + 1);
      return {
        mes: label,
        ocupacion: occupancyForMonth(ym),
        ocupacionPct: Math.round(occupancyPctForMonth(ym) * 10) / 10,
        ingresos: Math.round(incomeForMonth(ym)),
        gastosFijos: Math.round(fixedExpensesForMonth(ym)),
        gestion: Math.round(managementFeeForMonth(ym)),
        gastos: Math.round(totalExpensesForMonth(ym)),
        rentabilidad: Math.round(profitabilityForMonth(ym) * 10) / 10
      };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tenants, expenses, selYear]);

  const occupancyStats = useMemo(() => {
    const pcts = yearData.map(d => d.ocupacionPct);
    const roomMonths = yearData.reduce((s, d) => s + d.ocupacion, 0);
    const capacityRoomMonths = TOTAL_HABITACIONES * 12;
    return {
      mediaAnual: pcts.reduce((s, p) => s + p, 0) / 12,
      roomMonths,
      capacityRoomMonths,
      totalAnual: (roomMonths / capacityRoomMonths) * 100
    };
  }, [yearData]);

  const currentYm = ymKey(now.getFullYear(), now.getMonth() + 1);

  const alerts = useMemo(() => {
    const todayD = toDate(todayISO());
    const impagos = [];
    const finContrato = [];
    const fianzaNoCobrada = [];
    const fianzaNoDevuelta = [];

    tenants.forEach(t => {
      const end = effectiveEnd(t);
      const endD = toDate(end);
      const activeNow = t.activo && overlapsMonth(t, currentYm);

      if (activeNow && !(t.pagos && t.pagos[currentYm])) {
        impagos.push(t);
      }
      if (t.activo && endD) {
        const days = Math.round((endD - todayD) / 86400000);
        if (days >= 0 && days <= 30) {
          finContrato.push({ ...t, diasRestantes: days });
        }
      }
      if (t.activo && !t.fechaPagoFianza) {
        fianzaNoCobrada.push(t);
      }
      const contractEnded = (!t.activo) || (endD && endD < todayD);
      if (contractEnded && t.fechaPagoFianza && !t.fechaDevolucionFianza) {
        fianzaNoDevuelta.push(t);
      }
    });

    finContrato.sort((a, b) => a.diasRestantes - b.diasRestantes);
    return { impagos, finContrato, fianzaNoCobrada, fianzaNoDevuelta };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tenants]);

  const totalAlerts = alerts.impagos.length + alerts.finContrato.length + alerts.fianzaNoCobrada.length + alerts.fianzaNoDevuelta.length;

  const fianzas = useMemo(() => {
    let pendienteCobro = 0, enPoder = 0, devuelto = 0;
    tenants.forEach(t => {
      const importe = Number(t.fianzaImporte) || 0;
      if (importe <= 0) return;
      if (!t.fechaPagoFianza) pendienteCobro += importe;
      else if (!t.fechaDevolucionFianza) enPoder += importe;
      else devuelto += importe;
    });
    return { pendienteCobro, enPoder, devuelto };
  }, [tenants]);

  /* --------------------------- Acciones --------------------------- */

  function shiftMonth(delta) {
    let y = selYear, m = selMonthNum + delta;
    while (m > 12) { m -= 12; y += 1; }
    while (m < 1) { m += 12; y -= 1; }
    setSelectedMonth(ymKey(y, m));
  }

  function openNewTenant() {
    setEditingTenant(emptyTenant());
    setShowForm(true);
  }
  function openNewTenantForRoom(roomNumber) {
    setEditingTenant({ ...emptyTenant(), habitacion: String(roomNumber) });
    setShowForm(true);
  }
  function openEditTenant(t) {
    setEditingTenant({ ...t, pagos: { ...(t.pagos || {}) } });
    setShowForm(true);
  }
  function saveTenant(t) {
    if (!t.habitacion.trim() || !t.nombre.trim()) {
      notify("Indica al menos la habitación y el nombre.");
      return;
    }
    setTenants(prev => {
      const exists = prev.some(p => p.id === t.id);
      if (exists) return prev.map(p => (p.id === t.id ? t : p));
      return [...prev, t];
    });
    setShowForm(false);
    setEditingTenant(null);
    notify("Inquilino guardado.");
  }
  function deleteTenant(id) {
    setTenants(prev => prev.filter(p => p.id !== id));
    setConfirmDelete(null);
    notify("Inquilino eliminado.");
  }
  function freeRoom(t) {
    setTenants(prev => prev.map(p => (p.id === t.id ? { ...p, activo: false } : p)));
    setConfirmFree(null);
    notify(`Habitación ${t.habitacion} liberada. ${t.nombre} pasa al histórico.`);
  }
  function reactivateTenant(t) {
    setTenants(prev => prev.map(p => (p.id === t.id ? { ...p, activo: true } : p)));
    notify(`${t.nombre} vuelve a aparecer como inquilino activo.`);
  }
  function togglePagoQuick(tenantId, ym) {
    setTenants(prev => prev.map(t => {
      if (t.id !== tenantId) return t;
      const pagos = { ...(t.pagos || {}) };
      pagos[ym] = !pagos[ym];
      return { ...t, pagos };
    }));
  }
  function updateExpenseField(ym, field, value) {
    setExpenses(prev => ({ ...prev, [ym]: { ...(prev[ym] || {}), [field]: value } }));
  }
  function updateExpenseItems(ym, field, items) {
    setExpenses(prev => ({ ...prev, [ym]: { ...(prev[ym] || {}), [field]: items } }));
  }

  async function uploadFacturaGasto(ym, categoria, file) {
    try {
      const path = `gastos/${ym}/${categoria}/${Date.now()}_${sanitizeFileName(file.name)}`;
      const meta = await uploadDoc(path, file);
      setExpenses(prev => {
        const monthData = { ...(prev[ym] || {}) };
        monthData.facturas = { ...(monthData.facturas || {}), [categoria]: meta };
        return { ...prev, [ym]: monthData };
      });
      notify("Factura adjuntada.");
    } catch (e) {
      console.error("Error al subir factura", e);
      notify("No se pudo subir la factura. Comprueba tu conexión.");
    }
  }

  async function deleteFacturaGasto(ym, categoria) {
    const meta = expenses[ym]?.facturas?.[categoria];
    if (!meta) return;
    try {
      await deleteDoc(meta.path);
    } catch (e) {
      console.error("Error al borrar factura", e);
    }
    setExpenses(prev => {
      const monthData = { ...(prev[ym] || {}) };
      const facturas = { ...(monthData.facturas || {}) };
      delete facturas[categoria];
      monthData.facturas = facturas;
      return { ...prev, [ym]: monthData };
    });
    notify("Factura eliminada.");
  }

  async function viewDoc(path) {
    try {
      const url = await getDocSignedUrl(path);
      window.open(url, "_blank", "noopener,noreferrer");
    } catch (e) {
      console.error("Error al abrir documento", e);
      notify("No se pudo abrir el documento.");
    }
  }

  function handleImportClick() {
    fileInputRef.current && fileInputRef.current.click();
  }
  function handleImportFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const imported = rows
          .filter(r => String(r["Nombre"] || "").trim() || String(r["Habitacion"] || "").trim())
          .map(r => ({
            ...emptyTenant(),
            habitacion: String(r["Habitacion"] || "").trim(),
            nombre: String(r["Nombre"] || "").trim(),
            apellidos: String(r["Apellidos"] || "").trim(),
            nacionalidad: String(r["Nacionalidad"] || "").trim(),
            telefono: String(r["Telefono"] || "").trim(),
            correo: String(r["Correo"] || "").trim(),
            fechaInicio: parseExcelDate(r["Fecha Inicio Contrato"]) || todayISO(),
            fechaFin: parseExcelDate(r["Fecha Final Contrato"]),
            renovado: truthy(r["Renovado Contrato"]),
            nuevaFechaFin: parseExcelDate(r["Nueva Fecha Fin"]),
            empadronado: truthy(r["Empadronado"]),
            renta: Number(r["Renta"]) || 0,
            observaciones: String(r["Observaciones"] || "").trim()
          }));
        if (imported.length) {
          setTenants(prev => [...prev, ...imported]);
          notify(`Se importaron ${imported.length} inquilino(s). Revisa fianzas y pagos manualmente.`);
        } else {
          notify("No se encontraron filas de inquilinos en el archivo.");
        }
      } catch (err) {
        console.error(err);
        notify("No se pudo leer el archivo. Comprueba que sea el formato esperado.");
      }
      e.target.value = "";
    };
    reader.readAsArrayBuffer(file);
  }
  function handleExport() {
    const rows = tenants.map(t => ({
      "Habitacion": t.habitacion,
      "Nombre": t.nombre,
      "Apellidos": t.apellidos,
      "Nacionalidad": t.nacionalidad,
      "Telefono": t.telefono,
      "Correo": t.correo,
      "Fecha Inicio Contrato": t.fechaInicio,
      "Fecha Final Contrato": t.fechaFin,
      "Renovado Contrato": t.renovado ? "Si" : "No",
      "Nueva Fecha Fin": t.nuevaFechaFin,
      "Empadronado": t.empadronado ? "Si" : "No",
      "Renta": t.renta,
      "Fianza Importe": t.fianzaImporte,
      "Fecha Pago Fianza": t.fechaPagoFianza,
      "Fecha Devolucion Fianza": t.fechaDevolucionFianza,
      "Activo": t.activo ? "Si" : "No",
      "Observaciones": t.observaciones
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inquilinos");
    XLSX.writeFile(wb, `inquilinos_${todayISO()}.xlsx`);
  }

  function handleExportHistorico() {
    const rows = tenants
      .filter(t => !t.activo)
      .map(t => ({
        "Habitacion": t.habitacion,
        "Nombre": t.nombre,
        "Apellidos": t.apellidos,
        "Telefono": t.telefono,
        "Correo": t.correo,
        "Fecha Inicio Contrato": t.fechaInicio,
        "Fecha Final Contrato": t.fechaFin,
        "Hubo Prorroga": t.renovado ? "Si" : "No",
        "Fecha Prorroga (Nueva Fecha Fin)": t.nuevaFechaFin,
        "Renta Mensual": t.renta,
        "Empadronado": t.empadronado ? "Si" : "No",
        "Fecha Devolucion Fianza": t.fechaDevolucionFianza
      }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Historico");
    XLSX.writeFile(wb, `historico_inquilinos_${todayISO()}.xlsx`);
  }

  const monthLabel = `${MESES[selMonthNum - 1]} ${selYear}`;

  return (
    <div className="rg-root">
      <GlobalStyles />
      <div className="rg-demo-banner">
        <Sparkles size={13} />
        Estás viendo una <strong>demo con datos de ejemplo</strong> — nada de lo que hagas aquí se guarda de verdad.
      </div>
      <div className="rg-shell">
        {/* SIDEBAR */}
        <div className="rg-sidebar">
          <div className="rg-brand">
            <div className="rg-brand-mark"><KeyRound size={18} /></div>
            <div>
              <div className="rg-display" style={{ fontSize: 16, fontWeight: 700, lineHeight: 1 }}>Susalquia</div>
              <div style={{ fontSize: 9.5, color: "var(--text-dim)", letterSpacing: ".02em", lineHeight: 1.3, marginTop: 2 }}>La forma fácil de gestionar tus alquileres</div>
            </div>
          </div>

          <button className={`rg-nav-item ${view === "dashboard" ? "active" : ""}`} onClick={() => setView("dashboard")}>
            <Home size={16} /> Dashboard
          </button>
          <button className={`rg-nav-item ${view === "habitaciones" ? "active" : ""}`} onClick={() => setView("habitaciones")}>
            <LayoutGrid size={16} /> Habitaciones
          </button>
          <button className={`rg-nav-item ${view === "inquilinos" ? "active" : ""}`} onClick={() => setView("inquilinos")}>
            <Users size={16} /> Inquilinos
          </button>
          <button className={`rg-nav-item ${view === "gastos" ? "active" : ""}`} onClick={() => setView("gastos")}>
            <Wallet size={16} /> Gastos
          </button>
          <button className={`rg-nav-item ${view === "alertas" ? "active" : ""}`} onClick={() => setView("alertas")}>
            <AlertTriangle size={16} /> Alertas
            {totalAlerts > 0 && <span className="rg-nav-badge">{totalAlerts}</span>}
          </button>
          <button className={`rg-nav-item ${view === "historico" ? "active" : ""}`} onClick={() => setView("historico")}>
            <History size={16} /> Histórico
          </button>

          <div className="rg-sidebar-footer">
            <div className="rg-sidebar-footer-info">
              <span className="rg-save-indicator" style={{ color: "var(--accent)" }}><Sparkles size={12} /> Modo demo · datos de ejemplo</span>
              <div style={{ fontSize: 10.5, color: "var(--text-dim)", marginTop: 6, lineHeight: 1.5 }}>
                Nada se guarda de verdad. Prueba a editar, subir documentos o generar un contrato con total libertad.
              </div>
            </div>
            <button className="rg-icon-btn rg-logout-btn" onClick={reiniciarDemo} title="Reiniciar demo">
              <RotateCcw size={14} /> <span style={{ fontSize: 11 }}>Reiniciar demo</span>
            </button>
          </div>
        </div>

        {/* MAIN */}
        <div className="rg-main">
          {view === "dashboard" && (
            <DashboardView
              monthLabel={monthLabel} selectedMonth={selectedMonth} shiftMonth={shiftMonth}
              incomeForMonth={incomeForMonth} fixedExpensesForMonth={fixedExpensesForMonth}
              managementFeeForMonth={managementFeeForMonth} totalExpensesForMonth={totalExpensesForMonth}
              netProfitForMonth={netProfitForMonth} profitabilityForMonth={profitabilityForMonth}
              yearData={yearData} fianzas={fianzas} occupancyStats={occupancyStats}
              selYear={selYear} incomeEmpadronadosForYear={incomeEmpadronadosForYear}
              irpfReduccion={irpfReduccion} setIrpfReduccion={setIrpfReduccion}
            />
          )}

          {view === "habitaciones" && (
            <HabitacionesView
              tenants={tenants} selectedMonth={selectedMonth}
              onEdit={openEditTenant} onFree={setConfirmFree} onAddForRoom={openNewTenantForRoom}
            />
          )}

          {view === "inquilinos" && (
            <InquilinosView
              tenants={tenants} selectedMonth={selectedMonth} shiftMonth={shiftMonth} monthLabel={monthLabel}
              onNew={openNewTenant} onEdit={openEditTenant} onDelete={setConfirmDelete}
              onFree={setConfirmFree} onTogglePago={togglePagoQuick}
              onImportClick={handleImportClick} onExport={handleExport}
            />
          )}

          {view === "gastos" && (
            <GastosView
              selectedMonth={selectedMonth} monthLabel={monthLabel} shiftMonth={shiftMonth}
              expenses={expenses} updateExpenseField={updateExpenseField} updateExpenseItems={updateExpenseItems}
              incomeForMonth={incomeForMonth} fixedExpensesForMonth={fixedExpensesForMonth}
              managementFeeForMonth={managementFeeForMonth} totalExpensesForMonth={totalExpensesForMonth}
              netProfitForMonth={netProfitForMonth} yearData={yearData} selYear={selYear}
              onUploadFactura={uploadFacturaGasto} onDeleteFactura={deleteFacturaGasto} onViewDoc={viewDoc}
              updateGestionConfig={updateGestionConfig}
            />
          )}

          {view === "alertas" && (
            <AlertasView alerts={alerts} onSelectTenant={openEditTenant} />
          )}

          {view === "historico" && (
            <HistoricoView tenants={tenants} onEdit={openEditTenant} onExport={handleExportHistorico} onReactivate={reactivateTenant} />
          )}
        </div>
      </div>

      <button className="rg-mobile-logout" onClick={reiniciarDemo}>
        <RotateCcw size={15} /> Reiniciar demo
      </button>

      <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleImportFile} />

      {showForm && editingTenant && (
        <TenantFormModal
          tenant={editingTenant}
          onCancel={() => { setShowForm(false); setEditingTenant(null); }}
          onSave={saveTenant}
          notify={notify}
          onViewDoc={viewDoc}
        />
      )}

      {confirmDelete && (
        <div className="rg-modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div className="rg-modal rg-confirm-box" onClick={(e) => e.stopPropagation()}>
            <div className="rg-modal-title" style={{ marginBottom: 10 }}>¿Eliminar inquilino?</div>
            <p style={{ color: "var(--text-dim)", fontSize: 13.5, marginBottom: 18 }}>
              Se eliminará permanentemente a <strong style={{ color: "var(--text)" }}>{confirmDelete.nombre} {confirmDelete.apellidos}</strong> y su historial de pagos. Esta acción no se puede deshacer.
            </p>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button className="rg-btn rg-btn-ghost" onClick={() => setConfirmDelete(null)}>Cancelar</button>
              <button className="rg-btn rg-btn-danger" onClick={() => deleteTenant(confirmDelete.id)}>Eliminar</button>
            </div>
          </div>
        </div>
      )}

      {confirmFree && (
        <div className="rg-modal-overlay" onClick={() => setConfirmFree(null)}>
          <div className="rg-modal rg-confirm-box" onClick={(e) => e.stopPropagation()}>
            <div className="rg-modal-title" style={{ marginBottom: 10 }}>¿Liberar habitación {confirmFree.habitacion}?</div>
            <p style={{ color: "var(--text-dim)", fontSize: 13.5, marginBottom: 18 }}>
              <strong style={{ color: "var(--text)" }}>{confirmFree.nombre} {confirmFree.apellidos}</strong> desaparecerá del listado de inquilinos activos
              y pasará al <strong style={{ color: "var(--text)" }}>Histórico</strong>, con todos sus datos guardados. La habitación quedará libre
              para dar de alta a un nuevo inquilino con una ficha en blanco. Puedes deshacerlo desde el Histórico si lo necesitas.
            </p>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button className="rg-btn rg-btn-ghost" onClick={() => setConfirmFree(null)}>Cancelar</button>
              <button className="rg-btn" onClick={() => freeRoom(confirmFree)}>Liberar habitación</button>
            </div>
          </div>
        </div>
      )}

      {toast && <div className="rg-toast">{toast}</div>}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Selector de mes reutilizable                                        */
/* ------------------------------------------------------------------ */

function MonthPicker({ monthLabel, shiftMonth }) {
  return (
    <div className="rg-month-picker">
      <button className="rg-icon-btn" onClick={() => shiftMonth(-1)} aria-label="Mes anterior"><ChevronLeft size={16} /></button>
      <span className="rg-mono">{monthLabel}</span>
      <button className="rg-icon-btn" onClick={() => shiftMonth(1)} aria-label="Mes siguiente"><ChevronRight size={16} /></button>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Selector de fecha desplegable (día / mes / año, próximos 25 años)    */
/* ------------------------------------------------------------------ */

function DateField({ label, value, onChange }) {
  const now = new Date();
  const minYear = now.getFullYear() - 2;
  const maxYear = now.getFullYear() + 25;
  const years = [];
  for (let y = minYear; y <= maxYear; y++) years.push(y);

  const d = value ? toDate(value) : null;
  const day = d ? d.getDate() : "";
  const month = d ? d.getMonth() + 1 : "";
  const year = d ? d.getFullYear() : "";

  const dayCount = (month && year) ? daysInMonth(Number(year), Number(month)) : 31;
  const days = Array.from({ length: dayCount }, (_, i) => i + 1);

  function update(newDay, newMonth, newYear) {
    if (!newDay || !newMonth || !newYear) { onChange(""); return; }
    const maxDay = daysInMonth(newYear, newMonth);
    const safeDay = Math.min(newDay, maxDay);
    onChange(`${newYear}-${pad2(newMonth)}-${pad2(safeDay)}`);
  }

  return (
    <div className="rg-field">
      <label className="rg-label">{label}</label>
      <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
        <select
          className="rg-select"
          value={day}
          onChange={(e) => update(Number(e.target.value), month || now.getMonth() + 1, year || now.getFullYear())}
        >
          <option value="">Día</option>
          {days.map(dd => <option key={dd} value={dd}>{dd}</option>)}
        </select>
        <select
          className="rg-select"
          value={month}
          onChange={(e) => update(day || 1, Number(e.target.value), year || now.getFullYear())}
        >
          <option value="">Mes</option>
          {MESES.map((m, idx) => <option key={m} value={idx + 1}>{m}</option>)}
        </select>
        <select
          className="rg-select"
          value={year}
          onChange={(e) => update(day || 1, month || now.getMonth() + 1, Number(e.target.value))}
        >
          <option value="">Año</option>
          {years.map(y => <option key={y} value={y}>{y}</option>)}
        </select>
        {value && (
          <button type="button" className="rg-icon-btn" onClick={() => onChange("")} title="Borrar fecha">
            <X size={13} />
          </button>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Dashboard                                                            */
/* ------------------------------------------------------------------ */

function DashboardView({ monthLabel, selectedMonth, shiftMonth, incomeForMonth, fixedExpensesForMonth, managementFeeForMonth, totalExpensesForMonth, netProfitForMonth, profitabilityForMonth, yearData, fianzas, occupancyStats, selYear, incomeEmpadronadosForYear, irpfReduccion, setIrpfReduccion }) {
  const ingresos = incomeForMonth(selectedMonth);
  const gastosFijosMes = fixedExpensesForMonth(selectedMonth);
  const gestionMes = managementFeeForMonth(selectedMonth);
  const beneficio = netProfitForMonth(selectedMonth);
  const rentabilidad = profitabilityForMonth(selectedMonth);
  const ocupacionMesPct = occupancyForMonthPct(selectedMonth, yearData);
  const ingresosAcumulados = yearData.reduce((s, d) => s + d.ingresos, 0);
  const gastosFijosAcumulados = yearData.reduce((s, d) => s + d.gastosFijos, 0);
  const gestionAcumulada = yearData.reduce((s, d) => s + d.gestion, 0);
  const diferenciaAcumulada = ingresosAcumulados - gastosFijosAcumulados - gestionAcumulada;

  const tooltipStyle = { background: "#121834", border: "1px solid #262c4d", borderRadius: 8, fontSize: 12, color: "#e7e8f5" };

  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">Panorama general</h1>
          <div className="rg-sub">Resumen económico y de ocupación</div>
        </div>
        <MonthPicker monthLabel={monthLabel} shiftMonth={shiftMonth} />
      </div>

      <div className="rg-grid-cards cols-5">
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Ingresos cobrados</div>
          <div className="rg-stat-value" style={{ color: "var(--ok)" }}>{fmtMoney(ingresos)}</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Gastos (sin gestión)</div>
          <div className="rg-stat-value" style={{ color: "var(--danger)" }}>{fmtMoney(gastosFijosMes)}</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Gestión a pagar</div>
          <div className="rg-stat-value" style={{ color: "var(--warn)" }}>{fmtMoney(gestionMes)}</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Ingresos − Gastos</div>
          <div className="rg-stat-value" style={{ color: beneficio >= 0 ? "var(--ok)" : "var(--danger)" }}>{fmtMoney(beneficio)}</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Rentabilidad neta</div>
          <div className="rg-stat-value" style={{ color: "var(--accent)" }}>{rentabilidad.toFixed(1)}%</div>
        </div>
      </div>

      <div className="rg-grid-cards">
        <div className="rg-card rg-stat" style={{ borderColor: "var(--ok)" }}>
          <div className="rg-stat-label">Ingresos acumulados del año {selYear}</div>
          <div className="rg-stat-value" style={{ color: "var(--ok)" }}>{fmtMoney(ingresosAcumulados)}</div>
        </div>
        <div className="rg-card rg-stat" style={{ borderColor: "var(--danger)" }}>
          <div className="rg-stat-label">Gastos acumulados (sin gestión)</div>
          <div className="rg-stat-value" style={{ color: "var(--danger)" }}>{fmtMoney(gastosFijosAcumulados)}</div>
        </div>
        <div className="rg-card rg-stat" style={{ borderColor: "var(--warn)" }}>
          <div className="rg-stat-label">Gestión acumulada del año</div>
          <div className="rg-stat-value" style={{ color: "var(--warn)" }}>{fmtMoney(gestionAcumulada)}</div>
        </div>
        <div className="rg-card rg-stat" style={{ borderColor: "var(--accent)" }}>
          <div className="rg-stat-label">Diferencia acumulada del año</div>
          <div className="rg-stat-value" style={{ color: diferenciaAcumulada >= 0 ? "var(--ok)" : "var(--danger)" }}>{fmtMoney(diferenciaAcumulada)}</div>
        </div>
      </div>

      <div className="rg-grid-cards cols-3">
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Ocupación este mes ({TOTAL_HABITACIONES} hab.)</div>
          <div className="rg-stat-value" style={{ color: "var(--accent)" }}>{ocupacionMesPct.toFixed(1)}%</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Ocupación media anual {selYear}</div>
          <div className="rg-stat-value" style={{ color: "var(--accent)" }}>{occupancyStats.mediaAnual.toFixed(1)}%</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Ocupación total del año</div>
          <div className="rg-stat-value" style={{ color: "var(--accent)" }}>{occupancyStats.totalAnual.toFixed(1)}%</div>
          <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 4 }}>{occupancyStats.roomMonths} de {occupancyStats.capacityRoomMonths} habitación-mes</div>
        </div>
      </div>

      <div className="rg-grid-cards cols-3">
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Fianzas pendientes de cobro</div>
          <div className="rg-stat-value" style={{ color: "var(--info)" }}>{fmtMoney(fianzas.pendienteCobro)}</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Fianzas en poder (no devueltas)</div>
          <div className="rg-stat-value" style={{ color: "var(--warn)" }}>{fmtMoney(fianzas.enPoder)}</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Fianzas devueltas</div>
          <div className="rg-stat-value" style={{ color: "var(--text-dim)" }}>{fmtMoney(fianzas.devuelto)}</div>
        </div>
      </div>

      <div className="rg-card rg-chart-card">
        <div className="rg-chart-title">Ocupación por mes (%) · {selYear}</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={yearData} margin={{ top: 6, right: 16, left: -18, bottom: 0 }}>
            <CartesianGrid stroke="#262c4d" vertical={false} />
            <XAxis dataKey="mes" tick={{ fill: "#9498b8", fontSize: 11 }} axisLine={{ stroke: "#262c4d" }} tickLine={false} />
            <YAxis domain={[0, 100]} unit="%" tick={{ fill: "#9498b8", fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip contentStyle={tooltipStyle} cursor={{ fill: "#1a2142" }} formatter={(v, n, p) => [`${v}% (${p.payload.ocupacion}/${TOTAL_HABITACIONES} hab.)`, "Ocupación"]} />
            <Bar dataKey="ocupacionPct" name="Ocupación" fill="#6366f1" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="rg-card rg-chart-card">
        <div className="rg-chart-title">Ingresos vs. gastos · {selYear}</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={yearData} margin={{ top: 6, right: 16, left: -18, bottom: 0 }}>
            <CartesianGrid stroke="#262c4d" vertical={false} />
            <XAxis dataKey="mes" tick={{ fill: "#9498b8", fontSize: 11 }} axisLine={{ stroke: "#262c4d" }} tickLine={false} />
            <YAxis tick={{ fill: "#9498b8", fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip contentStyle={tooltipStyle} cursor={{ fill: "#1a2142" }} formatter={(v) => fmtMoney(v)} />
            <Bar dataKey="ingresos" name="Ingresos" fill="#22c55e" radius={[4, 4, 0, 0]} />
            <Bar dataKey="gastos" name="Gastos" fill="#ef4444" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="rg-card rg-chart-card">
        <div className="rg-chart-title">Rentabilidad neta (%) · {selYear}</div>
        <ResponsiveContainer width="100%" height={220}>
          <LineChart data={yearData} margin={{ top: 6, right: 16, left: -18, bottom: 0 }}>
            <CartesianGrid stroke="#262c4d" vertical={false} />
            <XAxis dataKey="mes" tick={{ fill: "#9498b8", fontSize: 11 }} axisLine={{ stroke: "#262c4d" }} tickLine={false} />
            <YAxis tick={{ fill: "#9498b8", fontSize: 11 }} axisLine={false} tickLine={false} unit="%" />
            <Tooltip contentStyle={tooltipStyle} formatter={(v) => `${v}%`} />
            <Line type="monotone" dataKey="rentabilidad" name="Rentabilidad" stroke="#2f6fed" strokeWidth={2.5} dot={{ r: 3, fill: "#2f6fed" }} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <IrpfEstimateCard
        selYear={selYear}
        incomeEmpadronadosForYear={incomeEmpadronadosForYear}
        irpfReduccion={irpfReduccion}
        setIrpfReduccion={setIrpfReduccion}
      />
    </>
  );
}

function occupancyForMonthPct(ym, yearData) {
  const mesIdx = Number(ym.split("-")[1]) - 1;
  const entry = yearData[mesIdx];
  return entry ? entry.ocupacionPct : 0;
}

/* ------------------------------------------------------------------ */
/* Estimación de IRPF sobre ingresos de inquilinos empadronados         */
/* ------------------------------------------------------------------ */

function IrpfEstimateCard({ selYear, incomeEmpadronadosForYear, irpfReduccion, setIrpfReduccion }) {
  const ingresosBrutos = incomeEmpadronadosForYear(selYear);
  const rendimientoReducido = ingresosBrutos * (1 - irpfReduccion);
  const cuotaEstimada = calcIrpfProgresivo(rendimientoReducido);
  const tipoMedio = rendimientoReducido > 0 ? (cuotaEstimada / rendimientoReducido) * 100 : 0;

  return (
    <div className="rg-card rg-chart-card" style={{ paddingBottom: 20 }}>
      <div className="rg-chart-title">Estimación de IRPF · Inquilinos empadronados · {selYear}</div>
      <div style={{ padding: "0 16px" }}>
        <div className="rg-field" style={{ maxWidth: 460 }}>
          <label className="rg-label">Reducción aplicable sobre el rendimiento (Art. 23.2 LIRPF)</label>
          <select className="rg-select" value={irpfReduccion} onChange={(e) => setIrpfReduccion(Number(e.target.value))}>
            {REDUCCIONES_ALQUILER.map(r => <option key={r.value} value={r.value}>{r.label}</option>)}
          </select>
        </div>

        <div className="rg-grid-cards cols-3" style={{ marginTop: 8, marginBottom: 8 }}>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">Ingresos declarables (empadronados)</div>
            <div className="rg-stat-value">{fmtMoney(ingresosBrutos)}</div>
          </div>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">Rendimiento neto reducido</div>
            <div className="rg-stat-value" style={{ color: "var(--accent)" }}>{fmtMoney(rendimientoReducido)}</div>
          </div>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">Cuota IRPF estimada</div>
            <div className="rg-stat-value" style={{ color: "var(--danger)" }}>{fmtMoney(cuotaEstimada)}</div>
          </div>
        </div>

        <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 4 }}>
          Tipo medio aplicado sobre el rendimiento reducido: <strong className="rg-mono" style={{ color: "var(--text)" }}>{tipoMedio.toFixed(1)}%</strong>
        </div>

        <div style={{ fontSize: 11.5, color: "var(--text-dim)", lineHeight: 1.6, marginTop: 12, borderTop: "1px solid var(--border)", paddingTop: 12 }}>
          <strong style={{ color: "var(--warn)" }}>Esto es solo una estimación orientativa</strong>, no una declaración de la renta:
          se calcula con la escala combinada (estatal + autonómica) de referencia del IRPF 2025/2026, aplicada como si este fuera tu único ingreso del año.
          En la práctica se suma al resto de tus rentas (trabajo, pensión, etc.) y tributa según tu tipo marginal real; además, el tipo autonómico
          exacto depende de tu Comunidad Autónoma. La reducción del alquiler de habitaciones solo aplica si constituye la vivienda habitual y
          permanente del inquilino, y legalmente todos los ingresos por alquiler deben declararse con independencia de si el inquilino está
          empadronado o no. Consulta con un asesor fiscal o gestor antes de presentar tu declaración.
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Inquilinos                                                           */
/* ------------------------------------------------------------------ */

function InquilinosView({ tenants, selectedMonth, shiftMonth, monthLabel, onNew, onEdit, onDelete, onFree, onTogglePago, onImportClick, onExport }) {
  const activos = tenants.filter(t => t.activo);
  const sorted = [...activos].sort((a, b) => String(a.habitacion).localeCompare(String(b.habitacion), "es", { numeric: true }));

  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">Inquilinos</h1>
          <div className="rg-sub">{activos.length} inquilino(s) activo(s) ahora mismo</div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <MonthPicker monthLabel={monthLabel} shiftMonth={shiftMonth} />
          <button className="rg-btn rg-btn-ghost" onClick={onImportClick}><Upload size={14} /> Importar</button>
          <button className="rg-btn rg-btn-ghost" onClick={onExport}><Download size={14} /> Exportar</button>
          <button className="rg-btn" onClick={onNew}><Plus size={15} /> Nuevo inquilino</button>
        </div>
      </div>

      <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 14 }}>
        💡 Para dar de alta un inquilino nuevo en una habitación concreta, o para liberar una habitación cuando alguien se va, te resultará más cómodo usar la sección <strong style={{ color: "var(--text)" }}>Habitaciones</strong>.
      </div>

      <div className="rg-card">
        {sorted.length === 0 ? (
          <div className="rg-empty">
            <Users size={30} />
            <div>No hay inquilinos activos ahora mismo.</div>
            <div style={{ fontSize: 12, marginTop: 4 }}>Añade uno nuevo, importa tu Excel, o revisa la sección Habitaciones.</div>
          </div>
        ) : (
          <div className="rg-table-wrap">
            <table className="rg-table">
              <thead>
                <tr>
                  <th>Hab.</th>
                  <th>Inquilino</th>
                  <th>Empadronado</th>
                  <th>Renta</th>
                  <th>Fin contrato</th>
                  <th>Pago {monthLabel.split(" ")[0]}</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {sorted.map(t => {
                  const end = effectiveEnd(t);
                  const paid = !!(t.pagos && t.pagos[selectedMonth]);
                  const overlapsThisMonth = overlapsMonth(t, selectedMonth);
                  return (
                    <tr key={t.id}>
                      <td className="rg-mono">{t.habitacion || "—"}</td>
                      <td>{t.nombre} {t.apellidos}</td>
                      <td>
                        {t.empadronado ? (
                          <span className="rg-badge rg-badge-ok">Sí</span>
                        ) : (
                          <span className="rg-badge rg-badge-neutral">No</span>
                        )}
                      </td>
                      <td className="rg-mono">{fmtMoney(t.renta)}</td>
                      <td>{fmtDate(end)}</td>
                      <td>
                        {overlapsThisMonth ? (
                          <span
                            className="rg-stamp"
                            style={{ color: paid ? "var(--ok)" : "var(--danger)" }}
                            onClick={() => onTogglePago(t.id, selectedMonth)}
                          >
                            {paid ? "Pagado" : "Pendiente"}
                          </span>
                        ) : (
                          <span style={{ color: "var(--text-dim)", fontSize: 12 }}>—</span>
                        )}
                      </td>
                      <td>
                        <div style={{ display: "flex", gap: 4 }}>
                          <button className="rg-icon-btn" onClick={() => onEdit(t)} title="Editar"><Pencil size={14} /></button>
                          <button className="rg-icon-btn" onClick={() => onFree(t)} title="Liberar habitación"><DoorOpen size={14} /></button>
                          <button className="rg-icon-btn" onClick={() => onDelete(t)} title="Eliminar"><Trash2 size={14} /></button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Habitaciones                                                         */
/* ------------------------------------------------------------------ */

function HabitacionesView({ tenants, selectedMonth, onEdit, onFree, onAddForRoom }) {
  const rooms = ROOM_LABELS;
  const ocupadas = rooms.filter(label => tenants.some(t => t.activo && matchesRoom(t.habitacion, label))).length;

  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">Habitaciones</h1>
          <div className="rg-sub">{ocupadas} de {TOTAL_HABITACIONES} ocupadas ahora mismo</div>
        </div>
      </div>

      <div className="rg-rooms-grid">
        {rooms.map(label => {
          const tenant = tenants.find(t => t.activo && matchesRoom(t.habitacion, label));
          const paid = tenant && !!(tenant.pagos && tenant.pagos[selectedMonth]);

          if (!tenant) {
            return (
              <div className="rg-room-card rg-room-free" key={label}>
                <div className="rg-room-number">Hab. {label}</div>
                <span className="rg-badge rg-badge-neutral" style={{ marginBottom: 10 }}><DoorOpen size={11} /> Libre</span>
                <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 14 }}>Sin inquilino actualmente.</div>
                <button className="rg-btn" style={{ width: "100%", justifyContent: "center" }} onClick={() => onAddForRoom(label)}>
                  <Plus size={15} /> Añadir inquilino
                </button>
              </div>
            );
          }

          return (
            <div className="rg-room-card rg-room-occupied" key={label}>
              <div className="rg-room-number">Hab. {label}</div>
              <span className="rg-badge rg-badge-ok" style={{ marginBottom: 10 }}><DoorClosed size={11} /> Ocupada</span>
              <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 2 }}>{tenant.nombre} {tenant.apellidos}</div>
              <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 10 }}>{fmtMoney(tenant.renta)}/mes · fin {fmtDate(effectiveEnd(tenant))}</div>
              <div style={{ marginBottom: 14 }}>
                <span className="rg-badge" style={{ background: paid ? "var(--ok-dim)" : "var(--danger-dim)", color: paid ? "var(--ok)" : "var(--danger)" }}>
                  {paid ? "Pago al día" : "Pago pendiente"}
                </span>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button className="rg-btn rg-btn-ghost" style={{ flex: 1, justifyContent: "center" }} onClick={() => onEdit(tenant)}>
                  <Pencil size={13} /> Ver ficha
                </button>
                <button className="rg-btn rg-btn-ghost" style={{ flex: 1, justifyContent: "center" }} onClick={() => onFree(tenant)}>
                  <DoorOpen size={13} /> Liberar
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Bloque de conceptos itemizados (Reparaciones / Otros gastos)         */
/* ------------------------------------------------------------------ */

function ItemsBlock({ selectedMonth, field, label, conceptos, addLabel, items, factura, onAddItem, onUpdateItem, onRemoveItem, onUploadFactura, onDeleteFactura, onViewDoc }) {
  const fileRef = useRef(null);
  const subtotal = items.reduce((s, it) => s + (Number(it.importe) || 0), 0);

  return (
    <div className="rg-items-block">
      <div className="rg-items-block-header">
        <span>{label}</span>
        <span className="rg-mono" style={{ fontSize: 12.5, color: "var(--text-dim)" }}>{fmtMoney(subtotal)}</span>
      </div>

      {items.length === 0 ? (
        <div className="rg-items-empty">Sin conceptos registrados este mes.</div>
      ) : (
        items.map(item => (
          <div className="rg-item-row" key={item.id}>
            <select
              className="rg-select" value={item.concepto}
              onChange={(e) => onUpdateItem(field, item.id, { concepto: e.target.value })}
            >
              {conceptos.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            <input
              type="number" min="0" step="0.01" className="rg-input rg-item-amount" placeholder="0"
              value={item.importe}
              onChange={(e) => onUpdateItem(field, item.id, { importe: e.target.value === "" ? "" : Number(e.target.value) })}
            />
            <button type="button" className="rg-icon-btn" onClick={() => onRemoveItem(field, item.id)} title="Eliminar">
              <Trash2 size={13} />
            </button>
          </div>
        ))
      )}

      <div className="rg-items-actions">
        <button type="button" className="rg-btn-link" onClick={() => onAddItem(field, conceptos[0])}>
          <Plus size={12} /> {addLabel}
        </button>

        {onUploadFactura && (
          <>
            <input
              type="file" accept=".pdf,image/*"
              ref={fileRef}
              style={{ display: "none" }}
              onChange={(e) => { if (e.target.files[0]) onUploadFactura(selectedMonth, field, e.target.files[0]); e.target.value = ""; }}
            />
            {factura ? (
              <div className="rg-factura-row" style={{ marginTop: 0 }}>
                <button type="button" className="rg-chip" onClick={() => onViewDoc(factura.path)} title={factura.name}>
                  <FileText size={11} /> Ver factura
                </button>
                <button type="button" className="rg-icon-btn" onClick={() => onDeleteFactura(selectedMonth, field)} title="Quitar factura">
                  <Trash2 size={12} />
                </button>
              </div>
            ) : (
              <button type="button" className="rg-btn-link" onClick={() => fileRef.current?.click()}>
                <Paperclip size={11} /> Adjuntar factura
              </button>
            )}
          </>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Gastos                                                               */
/* ------------------------------------------------------------------ */

function GastosView({ selectedMonth, monthLabel, shiftMonth, expenses, updateExpenseField, updateExpenseItems, incomeForMonth, fixedExpensesForMonth, managementFeeForMonth, totalExpensesForMonth, netProfitForMonth, yearData, selYear, onUploadFactura, onDeleteFactura, onViewDoc, updateGestionConfig }) {
  const monthExpenses = expenses[selectedMonth] || {};
  const ingresos = incomeForMonth(selectedMonth);
  const gestion = managementFeeForMonth(selectedMonth);
  const fijos = fixedExpensesForMonth(selectedMonth);
  const total = totalExpensesForMonth(selectedMonth);
  const beneficio = netProfitForMonth(selectedMonth);
  const facturaRefs = useRef({});
  const gestionConfig = expenses._settings || { modo: "porcentaje", porcentaje: 15, fijo: 0 };

  function addItem(field, defaultConcepto) {
    const current = monthExpenses[field] || [];
    updateExpenseItems(selectedMonth, field, [...current, { id: uid(), concepto: defaultConcepto, importe: "" }]);
  }
  function updateItem(field, id, patch) {
    const current = monthExpenses[field] || [];
    updateExpenseItems(selectedMonth, field, current.map(it => (it.id === id ? { ...it, ...patch } : it)));
  }
  function removeItem(field, id) {
    const current = monthExpenses[field] || [];
    updateExpenseItems(selectedMonth, field, current.filter(it => it.id !== id));
  }

  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">Gastos</h1>
          <div className="rg-sub">Registro mensual de gastos y configuración de la gestión</div>
        </div>
        <MonthPicker monthLabel={monthLabel} shiftMonth={shiftMonth} />
      </div>

      <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
        <div className="rg-section-title">Configuración de la gestión</div>
        <div className="rg-form-grid">
          <div className="rg-field">
            <label className="rg-label">Modelo de gestión</label>
            <select
              className="rg-select" value={gestionConfig.modo}
              onChange={(e) => updateGestionConfig({ modo: e.target.value })}
            >
              <option value="porcentaje">Porcentaje sobre rentas cobradas</option>
              <option value="fijo">Cantidad fija mensual</option>
            </select>
          </div>
          {gestionConfig.modo === "fijo" ? (
            <div className="rg-field">
              <label className="rg-label">Cantidad fija de gestión (€/mes)</label>
              <input
                type="number" min="0" step="0.01" className="rg-input"
                value={gestionConfig.fijo}
                onChange={(e) => updateGestionConfig({ fijo: e.target.value === "" ? "" : Number(e.target.value) })}
              />
            </div>
          ) : (
            <div className="rg-field">
              <label className="rg-label">Porcentaje de gestión (%)</label>
              <input
                type="number" min="0" max="100" step="0.5" className="rg-input"
                value={gestionConfig.porcentaje}
                onChange={(e) => updateGestionConfig({ porcentaje: e.target.value === "" ? "" : Number(e.target.value) })}
              />
            </div>
          )}
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)" }}>
          Este ajuste se aplica a todos los meses. Acordado entre la propiedad y quien gestiona el alquiler.
        </div>
      </div>

      <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
        <div className="rg-section-title">Gastos fijos de {monthLabel}</div>
        <div className="rg-form-grid">
          {GASTOS_FIJOS_KEYS.map(k => {
            const factura = monthExpenses.facturas?.[k];
            return (
              <div className="rg-field" key={k}>
                <label className="rg-label">{GASTOS_FIJOS_LABELS[k]} (€)</label>
                <input
                  type="number" min="0" step="0.01" className="rg-input"
                  value={monthExpenses[k] ?? ""}
                  onChange={(e) => updateExpenseField(selectedMonth, k, e.target.value === "" ? "" : Number(e.target.value))}
                  placeholder="0"
                />
                <input
                  type="file" accept=".pdf,image/*"
                  ref={(el) => { facturaRefs.current[k] = el; }}
                  style={{ display: "none" }}
                  onChange={(e) => { if (e.target.files[0]) onUploadFactura(selectedMonth, k, e.target.files[0]); e.target.value = ""; }}
                />
                {factura ? (
                  <div className="rg-factura-row">
                    <button type="button" className="rg-chip" onClick={() => onViewDoc(factura.path)} title={factura.name}>
                      <FileText size={11} /> Ver factura
                    </button>
                    <button type="button" className="rg-icon-btn" onClick={() => onDeleteFactura(selectedMonth, k)} title="Quitar factura">
                      <Trash2 size={12} />
                    </button>
                  </div>
                ) : (
                  <button type="button" className="rg-btn-link" onClick={() => facturaRefs.current[k]?.click()}>
                    <Paperclip size={11} /> Adjuntar factura
                  </button>
                )}
              </div>
            );
          })}
        </div>

        <div className="rg-section-title">Reparaciones</div>
        <ItemsBlock
          selectedMonth={selectedMonth} field="reparacionesItems" label="Conceptos de reparación"
          conceptos={REPARACIONES_CONCEPTOS} addLabel="Añadir reparación"
          items={monthExpenses.reparacionesItems || []} factura={monthExpenses.facturas?.reparacionesItems}
          onAddItem={addItem} onUpdateItem={updateItem} onRemoveItem={removeItem}
          onUploadFactura={onUploadFactura} onDeleteFactura={onDeleteFactura} onViewDoc={onViewDoc}
        />

        <div className="rg-section-title">Otros gastos</div>
        <ItemsBlock
          selectedMonth={selectedMonth} field="otrosItems" label="Otros conceptos"
          conceptos={OTROS_CONCEPTOS} addLabel="Añadir otro gasto"
          items={monthExpenses.otrosItems || []} factura={monthExpenses.facturas?.otrosItems}
          onAddItem={addItem} onUpdateItem={updateItem} onRemoveItem={removeItem}
          onUploadFactura={onUploadFactura} onDeleteFactura={onDeleteFactura} onViewDoc={onViewDoc}
        />

        <div className="rg-section-title">Resumen calculado</div>
        <div className="rg-grid-cards cols-3" style={{ marginBottom: 0 }}>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">Gastos totales (sin gestión)</div>
            <div className="rg-stat-value">{fmtMoney(fijos)}</div>
          </div>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">
              Gestión ({gestionConfig.modo === "fijo" ? "cantidad fija" : `${gestionConfig.porcentaje || 0}% s/ cobrado`})
            </div>
            <div className="rg-stat-value" style={{ color: "var(--accent)" }}>{fmtMoney(gestion)}</div>
          </div>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">Gastos totales</div>
            <div className="rg-stat-value" style={{ color: "var(--danger)" }}>{fmtMoney(total)}</div>
          </div>
        </div>
        <div style={{ marginTop: 12, fontSize: 12.5, color: "var(--text-dim)" }}>
          Ingresos cobrados este mes: <strong className="rg-mono" style={{ color: "var(--ok)" }}>{fmtMoney(ingresos)}</strong> · Beneficio neto: <strong className="rg-mono" style={{ color: beneficio >= 0 ? "var(--ok)" : "var(--danger)" }}>{fmtMoney(beneficio)}</strong>
        </div>
      </div>

      <div className="rg-card">
        <div style={{ padding: "16px 20px 4px 20px" }} className="rg-chart-title">Histórico {selYear}</div>
        <div className="rg-table-wrap">
          <table className="rg-table">
            <thead>
              <tr>
                <th>Mes</th>
                <th>Ingresos</th>
                <th>Gastos (sin gestión)</th>
                <th>Gestión</th>
                <th>Gastos totales</th>
                <th>Beneficio neto</th>
              </tr>
            </thead>
            <tbody>
              {MESES.map((m, idx) => {
                const ym = ymKey(selYear, idx + 1);
                const ing = incomeForMonth(ym);
                const fij = fixedExpensesForMonth(ym);
                const ges = managementFeeForMonth(ym);
                const tot = totalExpensesForMonth(ym);
                const ben = netProfitForMonth(ym);
                const isCurrent = ym === selectedMonth;
                return (
                  <tr key={ym} style={isCurrent ? { background: "var(--accent-dim)" } : undefined}>
                    <td>{m}</td>
                    <td className="rg-mono">{fmtMoney(ing)}</td>
                    <td className="rg-mono">{fmtMoney(fij)}</td>
                    <td className="rg-mono">{fmtMoney(ges)}</td>
                    <td className="rg-mono">{fmtMoney(tot)}</td>
                    <td className="rg-mono" style={{ color: ben >= 0 ? "var(--ok)" : "var(--danger)" }}>{fmtMoney(ben)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Alertas                                                              */
/* ------------------------------------------------------------------ */

function AlertGroup({ title, icon, color, items, renderItem, emptyText }) {
  return (
    <div className="rg-alert-group">
      <div className="rg-alert-group-header" style={{ color }}>{icon} {title} <span className="rg-mono" style={{ fontSize: 12, color: "var(--text-dim)" }}>({items.length})</span></div>
      {items.length === 0 ? (
        <div style={{ fontSize: 13, color: "var(--text-dim)", padding: "4px 2px 8px 2px" }}>{emptyText}</div>
      ) : (
        items.map(renderItem)
      )}
    </div>
  );
}

function AlertasView({ alerts, onSelectTenant }) {
  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">Alertas</h1>
          <div className="rg-sub">Impagos, contratos y fianzas que requieren atención</div>
        </div>
      </div>

      <AlertGroup
        title="Impagos del mes en curso" icon={<AlertTriangle size={17} />} color="var(--danger)"
        items={alerts.impagos} emptyText="No hay impagos pendientes este mes."
        renderItem={(t) => (
          <div className="rg-alert-item" key={t.id} onClick={() => onSelectTenant(t)}>
            <div>
              <div className="rg-alert-item-name">{t.nombre} {t.apellidos}</div>
              <div className="rg-alert-item-sub">Habitación {t.habitacion || "—"}</div>
            </div>
            <span className="rg-badge rg-badge-danger">{fmtMoney(t.renta)}</span>
          </div>
        )}
      />

      <AlertGroup
        title="Fin de contrato próximo (30 días)" icon={<AlertTriangle size={17} />} color="var(--warn)"
        items={alerts.finContrato} emptyText="No hay contratos venciendo en los próximos 30 días."
        renderItem={(t) => (
          <div className="rg-alert-item" key={t.id} onClick={() => onSelectTenant(t)}>
            <div>
              <div className="rg-alert-item-name">{t.nombre} {t.apellidos}</div>
              <div className="rg-alert-item-sub">Habitación {t.habitacion || "—"} · Termina el {fmtDate(effectiveEnd(t))}</div>
            </div>
            <span className="rg-badge rg-badge-warn">{t.diasRestantes === 0 ? "Hoy" : `${t.diasRestantes} días`}</span>
          </div>
        )}
      />

      <AlertGroup
        title="Fianza no cobrada" icon={<ShieldAlert size={17} />} color="var(--info)"
        items={alerts.fianzaNoCobrada} emptyText="Todas las fianzas de inquilinos activos están cobradas."
        renderItem={(t) => (
          <div className="rg-alert-item" key={t.id} onClick={() => onSelectTenant(t)}>
            <div>
              <div className="rg-alert-item-name">{t.nombre} {t.apellidos}</div>
              <div className="rg-alert-item-sub">Habitación {t.habitacion || "—"}</div>
            </div>
            <span className="rg-badge rg-badge-info">{fmtMoney(t.fianzaImporte)}</span>
          </div>
        )}
      />

      <AlertGroup
        title="Fianza no devuelta" icon={<ShieldCheck size={17} />} color="var(--accent)"
        items={alerts.fianzaNoDevuelta} emptyText="No hay fianzas pendientes de devolución."
        renderItem={(t) => (
          <div className="rg-alert-item" key={t.id} onClick={() => onSelectTenant(t)}>
            <div>
              <div className="rg-alert-item-name">{t.nombre} {t.apellidos}</div>
              <div className="rg-alert-item-sub">Habitación {t.habitacion || "—"} · Contrato finalizado el {fmtDate(effectiveEnd(t))}</div>
            </div>
            <span className="rg-badge" style={{ background: "var(--accent-dim)", color: "var(--accent)" }}>{fmtMoney(t.fianzaImporte)}</span>
          </div>
        )}
      />
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Histórico de inquilinos                                              */
/* ------------------------------------------------------------------ */

function HistoricoView({ tenants, onEdit, onExport, onReactivate }) {
  const historicos = tenants
    .filter(t => !t.activo)
    .sort((a, b) => {
      const da = toDate(effectiveEnd(a)) || toDate(a.fechaInicio) || new Date(0);
      const db = toDate(effectiveEnd(b)) || toDate(b.fechaInicio) || new Date(0);
      return db - da; // más recientes primero
    });

  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">Histórico de inquilinos</h1>
          <div className="rg-sub">{historicos.length} contrato(s) finalizado(s)</div>
        </div>
        <button className="rg-btn rg-btn-ghost" onClick={onExport}><Download size={14} /> Exportar histórico</button>
      </div>

      <div className="rg-card">
        {historicos.length === 0 ? (
          <div className="rg-empty">
            <History size={30} />
            <div>Todavía no hay inquilinos en el histórico.</div>
            <div style={{ fontSize: 12, marginTop: 4 }}>
              Cuando liberes una habitación desde la sección Habitaciones, ese inquilino aparecerá aquí.
            </div>
          </div>
        ) : (
          <div className="rg-table-wrap">
            <table className="rg-table">
              <thead>
                <tr>
                  <th>Hab.</th>
                  <th>Inquilino</th>
                  <th>Contacto</th>
                  <th>Inicio contrato</th>
                  <th>Fin contrato</th>
                  <th>Prórroga</th>
                  <th>Renta pagada</th>
                  <th>Empadronado</th>
                  <th>Devolución fianza</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {historicos.map(t => (
                  <tr key={t.id}>
                    <td className="rg-mono">{t.habitacion || "—"}</td>
                    <td>{t.nombre} {t.apellidos}</td>
                    <td style={{ fontSize: 12 }}>
                      {t.telefono || "—"}{t.telefono && t.correo ? " · " : ""}{t.correo || (!t.telefono ? "—" : "")}
                    </td>
                    <td>{fmtDate(t.fechaInicio)}</td>
                    <td>{fmtDate(effectiveEnd(t))}</td>
                    <td>
                      {t.renovado ? (
                        <span className="rg-badge rg-badge-info">Sí, hasta {fmtDate(t.nuevaFechaFin)}</span>
                      ) : (
                        <span className="rg-badge rg-badge-neutral">No</span>
                      )}
                    </td>
                    <td className="rg-mono">{fmtMoney(t.renta)}</td>
                    <td>
                      {t.empadronado ? (
                        <span className="rg-badge rg-badge-ok">Sí</span>
                      ) : (
                        <span className="rg-badge rg-badge-neutral">No</span>
                      )}
                    </td>
                    <td>
                      {t.fechaDevolucionFianza ? (
                        fmtDate(t.fechaDevolucionFianza)
                      ) : t.fianzaImporte > 0 ? (
                        <span className="rg-badge rg-badge-warn">Pendiente</span>
                      ) : (
                        "—"
                      )}
                    </td>
                    <td>
                      <div style={{ display: "flex", gap: 4 }}>
                        <button className="rg-icon-btn" onClick={() => onEdit(t)} title="Ver / editar"><Pencil size={14} /></button>
                        <button className="rg-icon-btn" onClick={() => onReactivate(t)} title="Reactivar (volver a inquilinos activos)"><RotateCcw size={14} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Formulario de inquilino (modal)                                      */
/* ------------------------------------------------------------------ */

function TenantFormModal({ tenant, onCancel, onSave, notify, onViewDoc }) {
  const [form, setForm] = useState(tenant);
  const [fianzaTouched, setFianzaTouched] = useState(!!tenant.fianzaImporte);
  const [uploadingContrato, setUploadingContrato] = useState(false);
  const [uploadingIdentidad, setUploadingIdentidad] = useState(false);
  const contratoInputRef = useRef(null);
  const identidadInputRef = useRef(null);
  const year = new Date().getFullYear();

  function set(field, value) {
    setForm(prev => ({ ...prev, [field]: value }));
  }
  function togglePago(ym) {
    setForm(prev => ({ ...prev, pagos: { ...(prev.pagos || {}), [ym]: !prev.pagos?.[ym] } }));
  }
  function setRenta(value) {
    setForm(prev => ({
      ...prev,
      renta: value,
      // Mientras la fianza no se haya editado a mano, sigue igualando una mensualidad
      fianzaImporte: fianzaTouched ? prev.fianzaImporte : value
    }));
  }
  function setFianza(value) {
    setFianzaTouched(true);
    set("fianzaImporte", value);
  }

  const [generandoContrato, setGenerandoContrato] = useState(false);
  async function handleGenerarContrato() {
    setGenerandoContrato(true);
    try {
      await generarContratoDocx(form);
      notify && notify("Contrato generado. Revisa tu carpeta de descargas.");
    } catch (e) {
      console.error("Error al generar contrato", e);
      notify && notify(e.message || "No se pudo generar el contrato.");
    } finally {
      setGenerandoContrato(false);
    }
  }

  async function handleUploadDocs(kind, fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    const field = kind === "contrato" ? "documentosContrato" : "documentosIdentidad";
    const setUploading = kind === "contrato" ? setUploadingContrato : setUploadingIdentidad;
    setUploading(true);
    try {
      const uploaded = [];
      for (const file of files) {
        const path = `tenants/${form.id}/${kind}/${Date.now()}_${sanitizeFileName(file.name)}`;
        const meta = await uploadDoc(path, file);
        uploaded.push(meta);
      }
      setForm(prev => ({ ...prev, [field]: [...(prev[field] || []), ...uploaded] }));
      notify && notify("Documento subido.");
    } catch (e) {
      console.error("Error al subir documento", e);
      notify && notify("No se pudo subir el documento. Comprueba tu conexión.");
    } finally {
      setUploading(false);
    }
  }

  async function handleDeleteDoc(kind, idx) {
    const field = kind === "contrato" ? "documentosContrato" : "documentosIdentidad";
    const doc = (form[field] || [])[idx];
    if (!doc) return;
    try {
      await deleteDoc(doc.path);
    } catch (e) {
      console.error("Error al borrar documento", e);
    }
    setForm(prev => ({ ...prev, [field]: (prev[field] || []).filter((_, i) => i !== idx) }));
  }

  return (
    <div className="rg-modal-overlay" onClick={onCancel}>
      <div className="rg-modal" onClick={(e) => e.stopPropagation()}>
        <div className="rg-modal-header">
          <div className="rg-modal-title">{tenant.nombre ? "Editar inquilino" : "Nuevo inquilino"}</div>
          <button className="rg-icon-btn" onClick={onCancel}><X size={18} /></button>
        </div>

        <div className="rg-section-title">Datos personales</div>
        <div className="rg-form-grid">
          <div className="rg-field">
            <label className="rg-label">Habitación</label>
            <input className="rg-input" value={form.habitacion} onChange={(e) => set("habitacion", e.target.value)} placeholder="Ej. 3" />
          </div>
          <div className="rg-field">
            <label className="rg-label">Nombre</label>
            <input className="rg-input" value={form.nombre} onChange={(e) => set("nombre", e.target.value)} />
          </div>
          <div className="rg-field">
            <label className="rg-label">Apellidos</label>
            <input className="rg-input" value={form.apellidos} onChange={(e) => set("apellidos", e.target.value)} />
          </div>
          <div className="rg-field">
            <label className="rg-label">Tipo de documento</label>
            <select className="rg-select" value={form.tipoDocumento} onChange={(e) => set("tipoDocumento", e.target.value)}>
              <option value="DNI">DNI</option>
              <option value="NIE">NIE</option>
              <option value="Pasaporte">Pasaporte</option>
            </select>
          </div>
          <div className="rg-field">
            <label className="rg-label">Número de documento</label>
            <input className="rg-input" value={form.numeroDocumento} onChange={(e) => set("numeroDocumento", e.target.value)} placeholder="Ej. 12345678A" />
          </div>
          <div className="rg-field">
            <label className="rg-label">Nacionalidad</label>
            <input className="rg-input" value={form.nacionalidad} onChange={(e) => set("nacionalidad", e.target.value)} />
          </div>
          <div className="rg-field">
            <label className="rg-label">Teléfono</label>
            <input className="rg-input" value={form.telefono} onChange={(e) => set("telefono", e.target.value)} />
          </div>
          <div className="rg-field">
            <label className="rg-label">Correo</label>
            <input className="rg-input" type="email" value={form.correo} onChange={(e) => set("correo", e.target.value)} />
          </div>
          <div className="rg-field">
            <label className="rg-check">
              <input type="checkbox" checked={!!form.empadronado} onChange={(e) => set("empadronado", e.target.checked)} />
              Empadronado
            </label>
          </div>
          <div className="rg-field">
            <label className="rg-check">
              <input type="checkbox" checked={!!form.activo} onChange={(e) => set("activo", e.target.checked)} />
              Habitación ocupada (activo)
            </label>
          </div>
        </div>

        <div className="rg-section-title">Contrato y renta</div>
        <div className="rg-form-grid">
          <DateField label="Fecha inicio contrato" value={form.fechaInicio} onChange={(v) => set("fechaInicio", v)} />
          <DateField label="Fecha final contrato" value={form.fechaFin} onChange={(v) => set("fechaFin", v)} />
          <div className="rg-field">
            <label className="rg-check">
              <input type="checkbox" checked={!!form.renovado} onChange={(e) => set("renovado", e.target.checked)} />
              Contrato renovado
            </label>
          </div>
          {form.renovado && (
            <DateField label="Nueva fecha fin" value={form.nuevaFechaFin} onChange={(v) => set("nuevaFechaFin", v)} />
          )}
          <div className="rg-field">
            <label className="rg-label">Renta mensual (€)</label>
            <input className="rg-input" type="number" min="0" step="0.01" value={form.renta} onChange={(e) => setRenta(e.target.value === "" ? "" : Number(e.target.value))} />
          </div>
        </div>

        <div className="rg-section-title">Fianza</div>
        <div className="rg-form-grid">
          <div className="rg-field">
            <label className="rg-label">Importe fianza (€)</label>
            <input className="rg-input" type="number" min="0" step="0.01" value={form.fianzaImporte} onChange={(e) => setFianza(Number(e.target.value))} />
            <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 4 }}>
              Por defecto, una mensualidad ({fmtMoney(form.renta)}). Puedes cambiarla libremente.
            </div>
          </div>
          <div />
          <DateField label="Fecha pago fianza" value={form.fechaPagoFianza} onChange={(v) => set("fechaPagoFianza", v)} />
          <DateField label="Fecha devolución fianza" value={form.fechaDevolucionFianza} onChange={(v) => set("fechaDevolucionFianza", v)} />
        </div>

        <div className="rg-section-title">Documentos</div>

        <div className="rg-doc-block">
          <div className="rg-doc-block-header">
            <span>Contrato y prórrogas</span>
            <div style={{ display: "flex", gap: 8 }}>
              <button type="button" className="rg-btn rg-btn-ghost" onClick={handleGenerarContrato} disabled={generandoContrato}>
                <FileSignature size={13} /> {generandoContrato ? "Generando…" : "Generar contrato"}
              </button>
              <button type="button" className="rg-btn rg-btn-ghost" onClick={() => contratoInputRef.current?.click()} disabled={uploadingContrato}>
                <Upload size={13} /> {uploadingContrato ? "Subiendo…" : "Subir archivo"}
              </button>
            </div>
            <input
              ref={contratoInputRef} type="file" accept=".pdf,image/*" multiple style={{ display: "none" }}
              onChange={(e) => { handleUploadDocs("contrato", e.target.files); e.target.value = ""; }}
            />
          </div>
          <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: -4, marginBottom: 10 }}>
            Genera el contrato en Word con los datos de esta ficha, listo para revisar, imprimir y firmar. La fecha del contrato se calcula automáticamente 3 días antes del inicio del alquiler.
          </div>
          {(form.documentosContrato || []).length === 0 ? (
            <div className="rg-doc-empty">Sin documentos todavía. Sube aquí el contrato inicial y, más adelante, cada prórroga como archivo aparte.</div>
          ) : (
            <ul className="rg-doc-list">
              {form.documentosContrato.map((d, i) => (
                <li key={d.path}>
                  <FileText size={13} />
                  <span className="rg-doc-name" onClick={() => onViewDoc(d.path)} title="Abrir documento">{d.name}</span>
                  <span className="rg-doc-date">{fmtDate(d.uploadedAt.slice(0, 10))}</span>
                  <button type="button" className="rg-icon-btn" onClick={() => handleDeleteDoc("contrato", i)} title="Eliminar"><Trash2 size={12} /></button>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="rg-doc-block">
          <div className="rg-doc-block-header">
            <span>Documento de identidad</span>
            <button type="button" className="rg-btn rg-btn-ghost" onClick={() => identidadInputRef.current?.click()} disabled={uploadingIdentidad}>
              <Upload size={13} /> {uploadingIdentidad ? "Subiendo…" : "Subir archivo"}
            </button>
            <input
              ref={identidadInputRef} type="file" accept=".pdf,image/*" multiple style={{ display: "none" }}
              onChange={(e) => { handleUploadDocs("identidad", e.target.files); e.target.value = ""; }}
            />
          </div>
          {(form.documentosIdentidad || []).length === 0 ? (
            <div className="rg-doc-empty">Sin documentos todavía. Sube aquí el DNI, NIE o pasaporte (anverso y reverso si hace falta).</div>
          ) : (
            <ul className="rg-doc-list">
              {form.documentosIdentidad.map((d, i) => (
                <li key={d.path}>
                  <FileText size={13} />
                  <span className="rg-doc-name" onClick={() => onViewDoc(d.path)} title="Abrir documento">{d.name}</span>
                  <span className="rg-doc-date">{fmtDate(d.uploadedAt.slice(0, 10))}</span>
                  <button type="button" className="rg-icon-btn" onClick={() => handleDeleteDoc("identidad", i)} title="Eliminar"><Trash2 size={12} /></button>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="rg-section-title">Pagos {year} (clic para marcar pagado / pendiente)</div>
        <div className="rg-pago-grid">
          {MESES_CORTOS.map((m, idx) => {
            const ym = ymKey(year, idx + 1);
            const paid = !!form.pagos?.[ym];
            return (
              <div key={ym} className={`rg-pago-cell ${paid ? "paid" : "unpaid"}`} onClick={() => togglePago(ym)}>
                {m}
              </div>
            );
          })}
        </div>

        <div className="rg-section-title">Observaciones</div>
        <textarea className="rg-textarea" value={form.observaciones} onChange={(e) => set("observaciones", e.target.value)} placeholder="Notas adicionales…" />

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 22 }}>
          <button className="rg-btn rg-btn-ghost" onClick={onCancel}>Cancelar</button>
          <button className="rg-btn" onClick={() => onSave({ ...form, renta: Number(form.renta) || 0, fianzaImporte: Number(form.fianzaImporte) || 0 })}>Guardar</button>
        </div>
      </div>
    </div>
  );
}
