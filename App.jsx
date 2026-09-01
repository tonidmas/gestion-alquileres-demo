import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from "recharts";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { CONTRATO_TEMPLATE_HABITACION_B64, CONTRATO_TEMPLATE_VIVIENDA_B64 } from "./contratoTemplate.js";
import {
  Home, Users, Wallet, AlertTriangle, KeyRound, Plus, X, Pencil, Trash2,
  Upload, Download, ShieldCheck, ShieldAlert, DoorOpen, DoorClosed,
  ChevronLeft, ChevronRight, Loader2, Check, RefreshCw, WifiOff, History,
  LayoutGrid, RotateCcw, FileText, Paperclip, LogOut, FileSignature, Settings,
  Eye, EyeOff, Wrench, Sparkles
} from "lucide-react";

/* ------------------------------------------------------------------ */
/* Utilidades                                                          */
/* ------------------------------------------------------------------ */

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const MESES_CORTOS = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
const GASTOS_FIJOS_KEYS = ["luz", "agua", "gas", "limpieza", "internet", "ibi", "comunidad"];
// De los gastos fijos, estos son los que tiene sentido repartir entre los inquilinos activos
// ese mes (suministros que varían con el uso) — el resto (limpieza, internet, IBI, comunidad)
// se consideran a cargo de la propiedad, no repartibles entre inquilinos.
const SUMINISTROS_REPARTIBLES_KEYS = ["luz", "agua", "gas"];

/* Días que un inquilino concreto estuvo activo dentro de un mes dado — se usa para repartir
   proporcionalmente los suministros de ese mes entre quienes estuvieron esos días. */
function diasActivoEnMes(t, year, month) {
  const dim = daysInMonth(year, month);
  const monthStartD = new Date(year, month - 1, 1);
  const monthEndD = new Date(year, month - 1, dim);
  const start = toDate(t.fechaInicio);
  const end = toDate(effectiveEndForOccupancy(t));
  if (!start) return 0;
  if (start > monthEndD) return 0;
  if (end && end < monthStartD) return 0;
  const rangeStart = start > monthStartD ? start : monthStartD;
  const rangeEnd = (end && end < monthEndD) ? end : monthEndD;
  if (rangeEnd < rangeStart) return 0;
  const days = Math.floor((rangeEnd - rangeStart) / 86400000) + 1;
  return Math.max(0, Math.min(dim, days));
}

/* Reparte el importe de los suministros de un mes entre los inquilinos que estuvieron activos,
   en proporción a los días que ocupó cada uno — no a partes iguales sin más, para que quien
   entró o se fue a mitad de mes pague solo su parte real. */
function repartoSuministros(tenants, expenses, ym) {
  const [year, month] = ym.split("-").map(Number);
  const mes = expenses[ym] || {};
  const totalSuministros = SUMINISTROS_REPARTIBLES_KEYS.reduce((s, k) => s + (Number(mes[k]) || 0), 0);

  const conDias = tenants
    .map(t => ({ tenant: t, dias: diasActivoEnMes(t, year, month) }))
    .filter(x => x.dias > 0);

  const totalDias = conDias.reduce((s, x) => s + x.dias, 0);

  const reparto = conDias.map(({ tenant, dias }) => ({
    tenant, dias,
    importe: totalDias > 0 ? Math.round((totalSuministros * dias / totalDias) * 100) / 100 : 0,
  }));

  return { totalSuministros, totalDias, reparto };
}

const GASTOS_FIJOS_LABELS = { luz: "Luz", agua: "Agua", gas: "Gas", limpieza: "Limpieza", internet: "Internet", ibi: "IBI", comunidad: "Comunidad de Propietarios" };
const REPARACIONES_CONCEPTOS = ["Electricidad", "Fontanería", "Albañilería", "Carpintería", "Pintura", "Electrodomésticos", "Muebles", "Ropa de Cama", "Cocina"];
const INCIDENCIA_ESTADOS = [
  { value: "abierta", label: "Abierta", color: "var(--danger)" },
  { value: "en_curso", label: "En curso", color: "var(--warn)" },
  { value: "cerrada", label: "Cerrada", color: "var(--ok)" },
];
const OTROS_CONCEPTOS = ["Notaría", "Registro de la Propiedad", "ITP", "Otros Impuestos"];

const DEFAULT_ROOM_LABELS = [];

/* Fecha real de inicio de la gestión de estos alquileres: el calendario anual de ocupación
   no marca nada como ocupado ni libre en los meses anteriores a esta fecha. */
const GESTION_INICIO = "2026-07-01";

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

/* La renta puede cambiar con el tiempo. En vez de un único valor fijo, cada inquilino guarda un
   historial (`historialRenta`): una lista de { desde: "YYYY-MM", importe }. Esta función devuelve
   cuál era la renta vigente en un mes concreto, buscando la entrada más reciente que ya hubiera
   entrado en vigor para ese mes — así, cambiar la renta en un mes no reescribe los meses
   anteriores, que conservan la que tenían entonces. */
function rentaEnMes(t, ym) {
  const historial = Array.isArray(t.historialRenta) ? t.historialRenta : [];
  if (historial.length === 0) return Number(t.renta) || 0;

  let vigente = null;
  for (const entrada of historial) {
    if (entrada.desde <= ym && (!vigente || entrada.desde > vigente.desde)) vigente = entrada;
  }
  if (vigente) return Number(vigente.importe) || 0;

  // El mes consultado es anterior a cualquier entrada del historial. Usamos la entrada más
  // ANTIGUA que tengamos como mejor referencia — nunca la más reciente, que causaría el mismo
  // fallo que estamos evitando (aplicar una subida de renta a meses ya pasados).
  const masAntigua = historial.reduce((min, e) => (!min || e.desde < min.desde ? e : min), null);
  return Number(masAntigua.importe) || 0;
}
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
/* Red de seguridad: un inquilino inactivo nunca puede estar ocupando la habitación hoy ni
   en el futuro, sea cual sea la fecha de fin que tenga guardada (por ejemplo, si se liberó
   la habitación con una versión anterior de la app y esa fecha se quedó desactualizada). */
function effectiveEndForOccupancy(t) {
  const end = effectiveEnd(t);
  if (!t.activo) {
    const today = todayISO();
    if (!end || end > today) return today;
  }
  return end;
}
function overlapsMonth(t, ym) {
  const [y, m] = ym.split("-").map(Number);
  const monthStart = new Date(y, m - 1, 1);
  const monthEnd = new Date(y, m, 0);
  const start = toDate(t.fechaInicio);
  const end = toDate(effectiveEndForOccupancy(t));
  if (start && start > monthEnd) return false;
  if (end && end < monthStart) return false;
  return true;
}
function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }
function daysInMonth(y, m) { return new Date(y, m, 0).getDate(); } // m: 1-12

/* Ocupación de una habitación en un mes concreto del año, en días y % (para el calendario anual) */
function roomMonthOccupancy(tenants, roomLabel, year, month) {
  const dim = daysInMonth(year, month);
  const monthStartD = new Date(year, month - 1, 1);
  const monthEndD = new Date(year, month - 1, dim);

  const gestionInicio = toDate(GESTION_INICIO);
  if (gestionInicio && monthEndD < gestionInicio) {
    return { percent: null, tooltip: `Antes del inicio de la gestión (${fmtDate(GESTION_INICIO)})` };
  }

  let totalDays = 0;
  const notes = [];

  tenants
    .filter(t => matchesRoom(t.habitacion, roomLabel))
    .forEach(t => {
      const start = toDate(t.fechaInicio);
      const end = toDate(effectiveEndForOccupancy(t));
      if (!start) return;
      if (start > monthEndD) return;
      if (end && end < monthStartD) return;

      const rangeStart = start > monthStartD ? start : monthStartD;
      const rangeEnd = (end && end < monthEndD) ? end : monthEndD;
      if (rangeEnd < rangeStart) return;

      const days = Math.floor((rangeEnd - rangeStart) / 86400000) + 1;
      totalDays += Math.max(0, Math.min(dim, days));

      const nombre = `${t.nombre || ""} ${t.apellidos || ""}`.trim() || "Inquilino";
      if (end && end >= monthStartD && end < monthEndD) {
        notes.push(`Fin de contrato: ${fmtDate(effectiveEndForOccupancy(t))} (${nombre})`);
      }
      if (start > monthStartD && start <= monthEndD) {
        notes.push(`Inicio de contrato: ${fmtDate(t.fechaInicio)} (${nombre})`);
      }
    });

  // Para pintar los colores en el orden correcto (rojo/verde según lo que pase primero
  // dentro del mes), miramos si el propio día 1 del mes está ocupado o no.
  const startsOccupied = tenants
    .filter(t => matchesRoom(t.habitacion, roomLabel))
    .some(t => {
      const start = toDate(t.fechaInicio);
      const end = toDate(effectiveEndForOccupancy(t));
      if (!start) return false;
      if (start > monthStartD) return false;
      if (end && end < monthStartD) return false;
      return true;
    });

  const percent = Math.max(0, Math.min(100, Math.round((totalDays / dim) * 100)));
  const tooltip = notes.length ? notes.join(" · ") : (percent > 0 ? "Ocupada todo el mes" : "Libre todo el mes");
  return { percent, tooltip, startsOccupied };
}

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

function validarContrato(tenant, arrendador, direccionVivienda) {
  const obligatorios = [];
  if (!tenant.nombre?.trim()) obligatorios.push("nombre");
  if (!tenant.apellidos?.trim()) obligatorios.push("apellidos");
  if (!tenant.numeroDocumento?.trim()) obligatorios.push("documento de identidad (DNI/NIE/Pasaporte)");
  if (!tenant.habitacion?.trim()) obligatorios.push("habitación");
  if (!tenant.fechaInicio) obligatorios.push("fecha de inicio de contrato");
  if (!tenant.fechaFin) obligatorios.push("fecha de fin de contrato");
  if (!tenant.renta) obligatorios.push("renta mensual");
  if (!tenant.fianzaImporte) obligatorios.push("importe de la fianza");
  if (!arrendador?.nombre?.trim()) obligatorios.push("nombre del arrendador/a (en Mi cuenta)");
  if (!arrendador?.documento?.trim()) obligatorios.push("documento del arrendador/a (en Mi cuenta)");
  if (!arrendador?.domicilio?.trim()) obligatorios.push("domicilio del arrendador/a (en Mi cuenta)");
  if (!direccionVivienda?.calle?.trim()) obligatorios.push("calle de la vivienda (en Mi cuenta)");
  if (!direccionVivienda?.numero?.trim()) obligatorios.push("número de la vivienda (en Mi cuenta)");
  if (!direccionVivienda?.cp?.trim()) obligatorios.push("código postal de la vivienda (en Mi cuenta)");
  if (!direccionVivienda?.localidad?.trim()) obligatorios.push("localidad de la vivienda (en Mi cuenta)");
  if (!direccionVivienda?.provincia?.trim()) obligatorios.push("provincia de la vivienda (en Mi cuenta)");
  if (!direccionVivienda?.refCatastral?.trim()) obligatorios.push("referencia catastral (en Mi cuenta)");
  if (!arrendador?.lugarFirma?.trim()) obligatorios.push("lugar de firma (en Mi cuenta)");

  // Datos de contacto del inquilino: importantes, pero no bloquean la generación — se puede
  // avisar y dejar que el propietario decida generar igualmente, dejando un hueco en blanco.
  const opcionales = [];
  if (!tenant.nacionalidad?.trim()) opcionales.push("nacionalidad");
  if (!tenant.telefono?.trim()) opcionales.push("teléfono");
  if (!tenant.correo?.trim()) opcionales.push("correo electrónico");

  return { obligatorios, opcionales };
}

const HUECO_EN_BLANCO = "______________________";
function valorOHueco(v) {
  return v?.trim() ? v.trim() : HUECO_EN_BLANCO;
}

function construirDireccionVivienda(direccion) {
  const partes = [`${direccion.calle.trim()}, número ${direccion.numero.trim()}`];
  const pisoLetra = [direccion.piso?.trim(), direccion.letra?.trim()].filter(Boolean).join(" ");
  if (pisoLetra) partes.push(pisoLetra);
  partes.push(`${direccion.cp.trim()} ${direccion.localidad.trim()} (${direccion.provincia.trim()})`);
  return partes.join(", ");
}

async function generarContratoDocx(tenant, arrendador, direccionVivienda, plantillaBytes, tipoUnidad) {
  const { obligatorios } = validarContrato(tenant, arrendador, direccionVivienda);
  if (obligatorios.length) {
    throw new Error(`Faltan datos para generar el contrato: ${obligatorios.join(", ")}.`);
  }

  const fechaFirmaISO = addDaysISO(tenant.fechaInicio, -3);
  const { dia, mes, anio } = fechaFirmaPartes(fechaFirmaISO);
  const renta = Number(tenant.renta) || 0;
  const fianza = Number(tenant.fianzaImporte) || 0;
  const meses = renta > 0 ? (Math.round((fianza / renta) * 10) / 10).toString().replace(".", ",") : "";

  const valores = {
    "@@NOMBRE_COMPLETO@@": `${tenant.nombre.trim()} ${tenant.apellidos.trim()}`,
    "@@NACIONALIDAD@@": valorOHueco(tenant.nacionalidad),
    "@@DOCUMENTO_NUM@@": tenant.numeroDocumento.trim(),
    "@@TELEFONO@@": valorOHueco(tenant.telefono),
    "@@CORREO@@": valorOHueco(tenant.correo),
    "@@HABITACION@@": tenant.habitacion.trim(),
    "@@FECHA_INICIO@@": fmtDate(tenant.fechaInicio),
    "@@FECHA_FIN@@": fmtDate(tenant.fechaFin),
    "@@RENTA_LETRAS@@": importeALetras(renta),
    "@@RENTA_CIFRA@@": renta.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    "@@FIANZA_LETRAS@@": importeALetras(fianza),
    "@@FIANZA_CIFRA@@": fianza.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    "@@FIANZA_MESES@@": meses,
    "@@LUGAR_FIRMA@@": arrendador.lugarFirma.trim(),
    "@@FECHA_FIRMA_DIA@@": dia,
    "@@FECHA_FIRMA_MES@@": mes,
    "@@FECHA_FIRMA_ANIO@@": anio,
    "@@ARRENDADOR_NOMBRE@@": arrendador.nombre.trim(),
    "@@ARRENDADOR_DOCUMENTO@@": arrendador.documento.trim(),
    "@@ARRENDADOR_DOMICILIO@@": arrendador.domicilio.trim(),
    "@@VIVIENDA_DIRECCION@@": construirDireccionVivienda(direccionVivienda),
    "@@REF_CATASTRAL@@": direccionVivienda.refCatastral.trim(),
  };

  const plantillaPorDefecto = tipoUnidad === "vivienda" ? CONTRATO_TEMPLATE_VIVIENDA_B64 : CONTRATO_TEMPLATE_HABITACION_B64;
  const bytes = plantillaBytes || base64ToUint8Array(plantillaPorDefecto);
  const zip = await JSZip.loadAsync(bytes);
  let xml = await zip.file("word/document.xml").async("text");
  for (const [token, valor] of Object.entries(valores)) {
    xml = xml.split(token).join(escaparXml(valor));
  }
  zip.file("word/document.xml", xml);

  // El pie de página también puede contener el marcador de dirección, si la plantilla lo usa
  const footerFile = Object.keys(zip.files).find(f => /word\/footer\d*\.xml$/.test(f));
  if (footerFile) {
    let footerXml = await zip.file(footerFile).async("text");
    let changed = false;
    for (const [token, valor] of Object.entries(valores)) {
      if (footerXml.includes(token)) { footerXml = footerXml.split(token).join(escaparXml(valor)); changed = true; }
    }
    if (changed) zip.file(footerFile, footerXml);
  }

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

/* Conexión a Supabase (base de datos en la nube, compartida entre dispositivos) */
/* ---------------------------------------------------------------------
   VERSIÓN DE DEMOSTRACIÓN: sin backend real. Los "documentos" se guardan
   solo en la memoria del navegador (como vista previa local), y todo lo
   demás vive en el estado de React — nada se envía a ningún servidor.
   Al recargar la página, o al pulsar "Reiniciar demo", todo vuelve a los
   datos de ejemplo originales.
--------------------------------------------------------------------- */
const DOCS_BUCKET = "documentos";
const demoFileStore = new Map();

function sanitizeFileName(name) {
  return String(name || "archivo").replace(/[^a-zA-Z0-9._-]/g, "_");
}

async function uploadDoc(path, file) {
  const url = URL.createObjectURL(file);
  demoFileStore.set(path, url);
  return { path, name: file.name, size: file.size, uploadedAt: new Date().toISOString() };
}

async function getDocSignedUrl(path) {
  const url = demoFileStore.get(path);
  if (!url) throw new Error("Documento no disponible en esta demo.");
  return url;
}

async function deleteDoc(path) {
  const url = demoFileStore.get(path);
  if (url) URL.revokeObjectURL(url);
  demoFileStore.delete(path);
}

async function deleteAllFilesUnderPrefix(prefix) {
  for (const key of Array.from(demoFileStore.keys())) {
    if (key.startsWith(prefix)) {
      URL.revokeObjectURL(demoFileStore.get(key));
      demoFileStore.delete(key);
    }
  }
}

function emptyTenant() {
  return {
    id: uid(),
    habitacion: "",
    vivienda: "",
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
    .rg-logout-btn {
      display: flex; align-items: center; justify-content: center; gap: 8px; width: 100%;
      margin-top: 14px; padding: 10px 14px; border-radius: 9px; font-size: 13px; font-weight: 600;
      background: var(--danger-dim); color: var(--danger); border: 1px solid transparent; cursor: pointer;
      transition: background .15s ease, border-color .15s ease;
    }
    .rg-logout-btn:hover { background: var(--danger); color: #fff; }

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

    /* Calendario anual de ocupación (Habitaciones): fija la columna de habitación, no la última */
    .rg-occupancy-table th, .rg-occupancy-table td { padding: 6px 5px; text-align: center; }
    .rg-occupancy-table th:first-child, .rg-occupancy-table td:first-child { text-align: left; padding-left: 12px; }
    .rg-occupancy-table th:last-child, .rg-occupancy-table td:last-child {
      position: static; box-shadow: none; background: transparent;
    }
    .rg-occupancy-table tr:hover td:last-child { background: var(--surface-alt); }
    .rg-occupancy-table th:first-child, .rg-occupancy-table td:first-child {
      position: sticky; left: 0; background: var(--surface); z-index: 1;
      box-shadow: 8px 0 10px -8px rgba(0,0,0,.35);
    }
    .rg-occupancy-table tr:hover td:first-child { background: var(--surface-alt); }
    .rg-occ-track { width: 100%; min-width: 34px; height: 20px; border-radius: 5px; background: var(--border); overflow: hidden; margin: 0 auto; display: flex; }
    .rg-occ-fill { height: 100%; transition: width .2s ease; }
    .rg-occ-fill-ocupado { background: var(--danger); }
    .rg-occ-fill-libre { background: var(--ok); }
    .rg-occ-track.rg-occ-na {
      background: repeating-linear-gradient(45deg, var(--surface-alt), var(--surface-alt) 4px, var(--bg) 4px, var(--bg) 8px);
      opacity: .6;
    }

    .rg-empty { padding: 40px 20px; text-align: center; color: var(--text-dim); }
    .rg-empty svg { margin-bottom: 10px; opacity: .5; }

    .rg-modal-overlay { position: fixed; inset: 0; background: rgba(8,10,13,.72); display: flex; align-items: flex-start; justify-content: center; z-index: 60; padding: 40px 16px; overflow-y: auto; }
    .rg-modal { background: var(--surface); border: 1px solid var(--border); border-radius: 14px; max-width: 760px; width: 100%; padding: 26px; margin-bottom: 40px; }
    .rg-modal-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 18px; }
    .rg-login-card { width: 100%; max-width: 360px; padding: 28px 26px; }
    .rg-link-btn { background: none; border: none; padding: 0; color: var(--accent); font-weight: 600; font-size: inherit; cursor: pointer; text-decoration: underline; }
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
    @media (max-width: 620px) {
      .rg-doc-block-header { flex-direction: column; align-items: stretch; gap: 8px; }
      .rg-doc-block-header > div { display: flex; flex-direction: column; gap: 8px; width: 100%; }
      .rg-doc-block-header > div button { width: 100%; justify-content: center; }
    }
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

/* Campo de contraseña con botón para mostrar/ocultar el texto */
function PasswordInput({ value, onChange, placeholder, autoComplete, autoFocus }) {
  const [show, setShow] = useState(false);
  return (
    <div style={{ position: "relative" }}>
      <input
        className="rg-input" type={show ? "text" : "password"}
        autoComplete={autoComplete} value={value} onChange={onChange} placeholder={placeholder}
        autoFocus={autoFocus} style={{ paddingRight: 38 }}
      />
      <button
        type="button" onClick={() => setShow(s => !s)} tabIndex={-1}
        title={show ? "Ocultar contraseña" : "Mostrar contraseña"}
        style={{
          position: "absolute", right: 6, top: "50%", transform: "translateY(-50%)",
          background: "none", border: "none", cursor: "pointer", color: "var(--text-dim)",
          padding: 6, display: "flex", alignItems: "center"
        }}
      >
        {show ? <EyeOff size={16} /> : <Eye size={16} />}
      </button>
    </div>
  );
}

function LoginScreen() {
  const [mode, setMode] = useState("login"); // "login" | "signup"
  const [nombre, setNombre] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [password2, setPassword2] = useState("");
  const [origenLead, setOrigenLead] = useState("");
  const [aceptaTerminos, setAceptaTerminos] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [submitting, setSubmitting] = useState(false);

  function switchMode(next) {
    setMode(next);
    setError("");
    setInfo("");
  }

  async function handleLogin(e) {
    e.preventDefault();
    if (!email.trim() || !password) {
      setError("Introduce tu correo y tu contraseña.");
      return;
    }
    setSubmitting(true);
    setError("");
    try {
      const { error: authError } = await supabase.auth.signInWithPassword({
        email: email.trim(),
        password
      });
      if (authError) {
        setError(authError.message === "Invalid login credentials"
          ? "Correo o contraseña incorrectos."
          : authError.message);
      }
      // Si el login es correcto, el listener onAuthStateChange de App
      // detecta la sesión y esta pantalla se sustituye automáticamente.
    } catch (e2) {
      console.error(e2);
      setError("No se pudo conectar. Comprueba tu conexión a internet.");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleSignup(e) {
    e.preventDefault();
    if (!nombre.trim()) { setError("Escribe tu nombre o el de tu negocio."); return; }
    if (!email.trim()) { setError("Introduce tu correo electrónico."); return; }
    if (password.length < 6) { setError("La contraseña debe tener al menos 6 caracteres."); return; }
    if (password !== password2) { setError("Las dos contraseñas no coinciden."); return; }
    if (!aceptaTerminos) { setError("Tienes que aceptar los Términos de Servicio y la Política de Privacidad para crear la cuenta."); return; }

    setSubmitting(true);
    setError("");
    setInfo("");
    try {
      const { data, error: authError } = await supabase.auth.signUp({
        email: email.trim(),
        password,
        options: { data: { nombre: nombre.trim(), origen_lead: origenLead || "web" } }
      });
      if (authError) {
        setError(authError.message === "User already registered"
          ? "Ya existe una cuenta con ese correo. Inicia sesión en vez de crear una nueva."
          : authError.message);
        return;
      }
      if (data?.session) {
        // Confirmación de correo desactivada en el proyecto: ya hay sesión iniciada,
        // el listener de App se encarga de cargar la app automáticamente.
        return;
      }
      // Confirmación de correo activada: hay que esperar a que confirmen desde su bandeja.
      setInfo("Cuenta creada. Revisa tu correo y confirma tu dirección para poder iniciar sesión.");
      setMode("login");
    } catch (e2) {
      console.error(e2);
      setError("No se pudo conectar. Comprueba tu conexión a internet.");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleResetRequest(e) {
    e.preventDefault();
    if (!email.trim()) { setError("Introduce tu correo electrónico."); return; }
    setSubmitting(true);
    setError("");
    setInfo("");
    try {
      const { error: authError } = await supabase.auth.resetPasswordForEmail(email.trim(), {
        redirectTo: window.location.origin
      });
      if (authError) { setError(authError.message); return; }
      // No confirmamos ni desmentimos si el correo existe en el sistema, por seguridad.
      setInfo("Si existe una cuenta con ese correo, te hemos enviado un enlace para restablecer tu contraseña. Revisa tu bandeja de entrada (y la carpeta de spam).");
    } catch (e2) {
      console.error(e2);
      setError("No se pudo conectar. Comprueba tu conexión a internet.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="rg-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", padding: 20 }}>
      <GlobalStyles />
      <form className="rg-card rg-login-card" onSubmit={mode === "login" ? handleLogin : mode === "signup" ? handleSignup : handleResetRequest}>
        <div className="rg-brand" style={{ marginBottom: 22, padding: 0 }}>
          <div className="rg-brand-mark"><KeyRound size={18} /></div>
          <div>
            <div className="rg-display" style={{ fontSize: 18, fontWeight: 700, lineHeight: 1 }}>Susalquia</div>
            <div style={{ fontSize: 10, color: "var(--text-dim)", marginTop: 2 }}>
              {mode === "reset" ? "Recupera el acceso a tu cuenta" : "La forma fácil de gestionar tus alquileres"}
            </div>
          </div>
        </div>

        {mode === "signup" && (
          <div className="rg-field">
            <label className="rg-label">Tu nombre o el de tu negocio</label>
            <input
              className="rg-input" value={nombre}
              onChange={(e) => setNombre(e.target.value)} placeholder="Ej. María López, o Inmuebles López"
              autoFocus
            />
          </div>
        )}

        <div className="rg-field">
          <label className="rg-label">Correo electrónico</label>
          <input
            className="rg-input" type="email" autoComplete="username" value={email}
            onChange={(e) => setEmail(e.target.value)} placeholder="tucorreo@ejemplo.com"
            autoFocus={mode === "login" || mode === "reset"}
          />
        </div>
        {mode !== "reset" && (
          <div className="rg-field">
            <label className="rg-label">Contraseña</label>
            <PasswordInput
              value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••"
              autoComplete={mode === "login" ? "current-password" : "new-password"}
            />
            {mode === "login" && (
              <div style={{ textAlign: "right", marginTop: 6 }}>
                <button type="button" onClick={() => switchMode("reset")} className="rg-link-btn" style={{ fontSize: 12 }}>
                  ¿Olvidaste tu contraseña?
                </button>
              </div>
            )}
          </div>
        )}
        {mode === "signup" && (
          <div className="rg-field">
            <label className="rg-label">Repite la contraseña</label>
            <PasswordInput
              value={password2} onChange={(e) => setPassword2(e.target.value)} placeholder="••••••••"
              autoComplete="new-password"
            />
          </div>
        )}

        {mode === "signup" && (
          <div className="rg-field">
            <label className="rg-label">¿Cómo nos has conocido? (opcional)</label>
            <select className="rg-select" value={origenLead} onChange={(e) => setOrigenLead(e.target.value)}>
              <option value="">Prefiero no decirlo</option>
              <option value="web">Buscando en internet</option>
              <option value="rrss">Redes sociales</option>
              <option value="recomendacion">Me lo ha recomendado alguien</option>
              <option value="otro">Otro</option>
            </select>
          </div>
        )}

        {mode === "signup" && (
          <label style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 12, color: "var(--text-dim)", marginBottom: 14, cursor: "pointer", lineHeight: 1.5 }}>
            <input
              type="checkbox" checked={aceptaTerminos}
              onChange={(e) => setAceptaTerminos(e.target.checked)}
              style={{ marginTop: 2, flexShrink: 0 }}
            />
            <span>
              He leído y acepto los{" "}
              <a href="https://susalquia.com/terminos.html" target="_blank" rel="noopener noreferrer" style={{ color: "var(--accent)" }}>
                Términos de Servicio
              </a>{" "}
              y la{" "}
              <a href="https://susalquia.com/privacidad.html" target="_blank" rel="noopener noreferrer" style={{ color: "var(--accent)" }}>
                Política de Privacidad
              </a>
              .
            </span>
          </label>
        )}

        {error && (
          <div style={{ background: "var(--danger-dim)", color: "var(--danger)", borderRadius: 8, padding: "9px 12px", fontSize: 12.5, marginBottom: 14 }}>{error}</div>
        )}
        {info && (
          <div style={{ background: "var(--ok-dim)", color: "var(--ok)", borderRadius: 8, padding: "9px 12px", fontSize: 12.5, marginBottom: 14 }}>{info}</div>
        )}

        <button type="submit" className="rg-btn" style={{ width: "100%", justifyContent: "center" }} disabled={submitting}>
          {submitting
            ? (mode === "login" ? "Entrando…" : mode === "signup" ? "Creando cuenta…" : "Enviando…")
            : (mode === "login" ? "Iniciar sesión" : mode === "signup" ? "Crear cuenta gratis" : "Enviar enlace de recuperación")}
        </button>

        {mode === "login" && (
          <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginTop: 16, textAlign: "center", lineHeight: 1.5 }}>
            ¿Todavía no tienes cuenta?{" "}
            <button type="button" onClick={() => switchMode("signup")} className="rg-link-btn">Crear una gratis</button>
          </div>
        )}
        {mode === "signup" && (
          <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginTop: 16, textAlign: "center", lineHeight: 1.5 }}>
            ¿Ya tienes cuenta?{" "}
            <button type="button" onClick={() => switchMode("login")} className="rg-link-btn">Iniciar sesión</button>
          </div>
        )}
        {mode === "reset" && (
          <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginTop: 16, textAlign: "center", lineHeight: 1.5 }}>
            <button type="button" onClick={() => switchMode("login")} className="rg-link-btn">← Volver a iniciar sesión</button>
          </div>
        )}
      </form>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Pantalla para establecer una contraseña nueva                        */
/* (aparece al volver del enlace de recuperación recibido por correo)   */
/* ------------------------------------------------------------------ */

function ResetPasswordScreen({ onDone }) {
  const [password, setPassword] = useState("");
  const [password2, setPassword2] = useState("");
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    if (password.length < 6) { setError("La contraseña debe tener al menos 6 caracteres."); return; }
    if (password !== password2) { setError("Las dos contraseñas no coinciden."); return; }
    setSubmitting(true);
    setError("");
    try {
      const { error: authError } = await supabase.auth.updateUser({ password });
      if (authError) { setError(authError.message); return; }
      onDone();
    } catch (e2) {
      console.error(e2);
      setError("No se pudo conectar. Comprueba tu conexión a internet.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="rg-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", padding: 20 }}>
      <GlobalStyles />
      <form className="rg-card rg-login-card" onSubmit={handleSubmit}>
        <div className="rg-brand" style={{ marginBottom: 22, padding: 0 }}>
          <div className="rg-brand-mark"><KeyRound size={18} /></div>
          <div>
            <div className="rg-display" style={{ fontSize: 18, fontWeight: 700, lineHeight: 1 }}>Susalquia</div>
            <div style={{ fontSize: 10, color: "var(--text-dim)", marginTop: 2 }}>Elige tu nueva contraseña</div>
          </div>
        </div>

        <div className="rg-field">
          <label className="rg-label">Nueva contraseña</label>
          <PasswordInput
            value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••"
            autoComplete="new-password" autoFocus
          />
        </div>
        <div className="rg-field">
          <label className="rg-label">Repite la nueva contraseña</label>
          <PasswordInput
            value={password2} onChange={(e) => setPassword2(e.target.value)} placeholder="••••••••"
            autoComplete="new-password"
          />
        </div>

        {error && (
          <div style={{ background: "var(--danger-dim)", color: "var(--danger)", borderRadius: 8, padding: "9px 12px", fontSize: 12.5, marginBottom: 14 }}>{error}</div>
        )}

        <button type="submit" className="rg-btn" style={{ width: "100%", justifyContent: "center" }} disabled={submitting}>
          {submitting ? "Guardando…" : "Guardar nueva contraseña"}
        </button>
      </form>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* App                                                                  */
/* ------------------------------------------------------------------ */

/* ---------------------------------------------------------------------
   Datos de ejemplo (MODO DEMO)
   La demo arranca completamente vacía a propósito — así la guía de
   "Primeros pasos" tiene sentido de verdad, en vez de encontrarse todo
   ya hecho de antemano.
--------------------------------------------------------------------- */
const DEMO_MONTH = "2026-08";

function buildDemoRoomLabels() {
  return [];
}

function buildDemoTenants() {
  return [];
}

function buildDemoIncidencias() {
  return [];
}

function buildDemoExpenses() {
  return {
    _settings: { modo: "porcentaje", porcentaje: 15, fijo: 0 },
  };
}

function AppInner() {
  const now = todayDate();
  const [tenants, setTenants] = useState(() => buildDemoTenants());
  const [incidencias, setIncidencias] = useState(() => buildDemoIncidencias());
  const [expenses, setExpenses] = useState(() => buildDemoExpenses());
  const [view, setView] = useState("dashboard");
  const [loaded] = useState(true);
  const [saving] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState(DEMO_MONTH);
  const [showForm, setShowForm] = useState(false);
  const [editingTenant, setEditingTenant] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [undoBuffer, setUndoBuffer] = useState(null);
  const [confirmFree, setConfirmFree] = useState(null);
  const [freeDate, setFreeDate] = useState(todayISO());
  function openConfirmFree(t) {
    setFreeDate(todayISO());
    setConfirmFree(t);
  }
  const [toast, setToast] = useState(null);
  const [irpfReduccion, setIrpfReduccion] = useState(0.5);
  const [roomLabels, setRoomLabels] = useState(() => buildDemoRoomLabels());
  const [connError] = useState(false);
  const [accountId] = useState("demo");
  const [accountPlan] = useState({ plan: "gestor", max_unidades: 999, max_usuarios: 99 });
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
    setIncidencias(buildDemoIncidencias());
    setExpenses(buildDemoExpenses());
    setRoomLabels(buildDemoRoomLabels());
    setIrpfReduccion(0.5);
    setSelectedMonth(DEMO_MONTH);
    setView("dashboard");
    notify("Demo reiniciada, sin ningún dato — como al entrar la primera vez.");
  }
  const handleSignOut = reiniciarDemo;

  function exportarCopiaSeguridad() {
    const backup = {
      exportado_en: new Date().toISOString(),
      cuenta: { correo: "demo@susalquia.com", plan: accountPlan.plan },
      habitaciones: roomLabels,
      inquilinos: tenants,
      gastos: expenses,
      irpf_reduccion: irpfReduccion,
      arrendador: getArrendadorConfig(),
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `susalquia_copia_seguridad_${todayISO()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    notify("Copia de seguridad descargada.");
  }

  /* En la demo no hay cuenta real que eliminar ni backend donde guardar — "reiniciar" es el
     equivalente exacto, y ya está definido más arriba (reiniciarDemo). */


  /* --------------------------- Cálculos --------------------------- */

  function incomeForMonth(ym) {
    return tenants.reduce((sum, t) => sum + (t.pagos && t.pagos[ym] ? rentaEnMes(t, ym) : 0), 0);
  }
  function occupancyForMonth(ym) {
    return tenants.filter(t => overlapsMonth(t, ym)).length;
  }
  function occupancyPctForMonth(ym) {
    if (roomLabels.length === 0) return 0;
    return (occupancyForMonth(ym) / roomLabels.length) * 100;
  }
  function incomeForYear(year) {
    let total = 0;
    for (let m = 1; m <= 12; m++) total += incomeForMonth(ymKey(year, m));
    return total;
  }
  function deductibleExpensesForYear(year) {
    let total = 0;
    for (let m = 1; m <= 12; m++) total += totalExpensesForMonth(ymKey(year, m));
    return total;
  }
  /* Varias viviendas en la misma cuenta: cada gasto fijo se guarda como { "": importe } (formato
     de siempre, "General") o, si hay más de una vivienda en uso, como un objeto con un importe
     por cada vivienda: { "": 10, "Piso Alicante": 40 }. Las cuentas con una sola vivienda o solo
     habitaciones sueltas nunca ven ningún cambio — siguen usando un único número, como siempre. */
  /* Cada unidad (habitación o vivienda completa) declara su tipo al crearse — así el
     vocabulario de toda la app y la plantilla de contrato usada se adaptan solas, sin que el
     propietario tenga que "acordarse" de nada después. Las unidades ya existentes de antes de
     este cambio, sin tipo registrado, se tratan como "habitación" (su comportamiento de
     siempre, sin ninguna sorpresa). */
  function getUnitType(label) {
    return expenses._settings?.unitTypes?.[label] || "habitacion";
  }
  function setUnitType(label, tipo) {
    setExpenses(prev => ({
      ...prev,
      _settings: { ...(prev._settings || {}), unitTypes: { ...(prev._settings?.unitTypes || {}), [label]: tipo } },
    }));
  }
  /* Para el checklist de bienvenida: si ya se rellenó al menos una dirección de vivienda, si
     alguna vez se generó un contrato, y si hay algún gasto introducido este mes. */
  function direccionViviendaRellena() {
    const dir = getDireccionVivienda("");
    return !!(dir && dir.calle && dir.calle.trim());
  }
  function marcarContratoGenerado() {
    if (!expenses._settings?.contratoGeneradoAlgunaVez) {
      updateArrendadorConfig({ contratoGeneradoAlgunaVez: true });
    }
  }
  function hayGastosIntroducidos() {
    const mes = expenses[selectedMonth] || {};
    const algunFijo = GASTOS_FIJOS_KEYS.some(k => gastoFijoTotal(mes[k]) > 0);
    const algunItem = (mes.reparacionesItems?.length > 0) || (mes.otrosItems?.length > 0);
    return algunFijo || algunItem;
  }
  function listaViviendas() {
    const set = new Set();
    tenants.forEach(t => { if (t.vivienda && t.vivienda.trim()) set.add(t.vivienda.trim()); });
    // Cualquier unidad marcada como "vivienda completa" es, ella sola, su propio grupo — sin
    // necesidad de que nadie escriba una etiqueta a mano.
    roomLabels.forEach(label => { if (getUnitType(label) === "vivienda") set.add(label); });
    return Array.from(set).sort((a, b) => a.localeCompare(b, "es"));
  }
  /* El nombre de la sección se adapta a lo que el cliente tiene configurado — nadie que solo
     gestione viviendas completas debería tener que leer "Habitaciones" por toda la app. */
  function etiquetaUnidades() {
    if (roomLabels.length === 0) return "Habitaciones";
    const hayHabitaciones = roomLabels.some(label => getUnitType(label) !== "vivienda");
    const hayViviendas = roomLabels.some(label => getUnitType(label) === "vivienda");
    if (hayViviendas && !hayHabitaciones) return "Viviendas";
    if (hayViviendas && hayHabitaciones) return "Habitaciones y viviendas";
    return "Habitaciones";
  }
  function gastoFijoValor(entry, viviendaKey) {
    if (entry == null) return "";
    if (typeof entry === "object") return entry[viviendaKey] ?? "";
    return viviendaKey === "" ? entry : "";
  }
  function gastoFijoPorVivienda(entry, viviendaKey) {
    if (entry == null) return 0;
    if (typeof entry === "object") return Number(entry[viviendaKey]) || 0;
    return viviendaKey === "" ? (Number(entry) || 0) : 0;
  }
  function setGastoFijoValor(prevEntry, viviendaKey, value) {
    const base = (prevEntry != null && typeof prevEntry === "object") ? { ...prevEntry } : (prevEntry != null ? { "": prevEntry } : {});
    base[viviendaKey] = value;
    return base;
  }
  function updateExpenseFieldVivienda(ym, field, viviendaKey, value) {
    setExpenses(prev => {
      const monthData = prev[ym] || {};
      const nuevoValor = setGastoFijoValor(monthData[field], viviendaKey, value);
      return { ...prev, [ym]: { ...monthData, [field]: nuevoValor } };
    });
  }
  function viviendaEfectivaDeTenant(t) {
    if (t.vivienda && t.vivienda.trim()) return t.vivienda.trim();
    if (getUnitType(t.habitacion) === "vivienda") return t.habitacion;
    return "";
  }
  function incomeForMonthPorVivienda(ym, viviendaKey) {
    return tenants.reduce((sum, t) => {
      if (viviendaEfectivaDeTenant(t) !== viviendaKey) return sum;
      return sum + (t.pagos && t.pagos[ym] ? rentaEnMes(t, ym) : 0);
    }, 0);
  }
  function fixedExpensesForMonthPorVivienda(ym, viviendaKey) {
    const e = expenses[ym] || {};
    const simple = GASTOS_FIJOS_KEYS.reduce((s, k) => s + gastoFijoPorVivienda(e[k], viviendaKey), 0);
    const reparaciones = (e.reparacionesItems || []).filter(it => (it.vivienda || "") === viviendaKey).reduce((s, it) => s + (Number(it.importe) || 0), 0);
    const otros = (e.otrosItems || []).filter(it => (it.vivienda || "") === viviendaKey).reduce((s, it) => s + (Number(it.importe) || 0), 0);
    return simple + reparaciones + otros;
  }
  function managementFeeForMonthPorVivienda(ym, viviendaKey) {
    const cfg = getGestionConfig();
    const ingresoTotal = incomeForMonth(ym);
    if (cfg.modo === "fijo") {
      if (ingresoTotal <= 0) return 0;
      return (Number(cfg.fijo) || 0) * (incomeForMonthPorVivienda(ym, viviendaKey) / ingresoTotal);
    }
    return incomeForMonthPorVivienda(ym, viviendaKey) * ((Number(cfg.porcentaje) || 0) / 100);
  }
  function netProfitForMonthPorVivienda(ym, viviendaKey) {
    return incomeForMonthPorVivienda(ym, viviendaKey) - fixedExpensesForMonthPorVivienda(ym, viviendaKey) - managementFeeForMonthPorVivienda(ym, viviendaKey);
  }
  function profitabilityForMonthPorVivienda(ym, viviendaKey) {
    const inc = incomeForMonthPorVivienda(ym, viviendaKey);
    if (inc <= 0) return 0;
    return (netProfitForMonthPorVivienda(ym, viviendaKey) / inc) * 100;
  }
  /* Desde cuándo se cuenta el "acumulado" de una vivienda: la fecha de inicio del inquilino
     activo más antiguo que ocupe esa vivienda ahora mismo. Si la vivienda no tiene inquilino
     activo, no hay nada que acumular todavía. */
  function acumuladoDesdeInicioVivienda(viviendaKey) {
    const activos = tenants.filter(t => t.activo && viviendaEfectivaDeTenant(t) === viviendaKey);
    if (activos.length === 0) return null;
    let fechaInicio = null;
    activos.forEach(t => {
      const d = toDate(t.fechaInicio);
      if (d && (!fechaInicio || d < fechaInicio)) fechaInicio = d;
    });
    if (!fechaInicio) return null;

    const [selY, selM] = selectedMonth.split("-").map(Number);
    let ingresos = 0, gastos = 0;
    let cursorAnio = fechaInicio.getFullYear();
    let cursorMes = fechaInicio.getMonth() + 1;
    // Recorre mes a mes desde el inicio del contrato hasta el mes que se está viendo
    while (cursorAnio < selY || (cursorAnio === selY && cursorMes <= selM)) {
      const ym = ymKey(cursorAnio, cursorMes);
      ingresos += incomeForMonthPorVivienda(ym, viviendaKey);
      gastos += fixedExpensesForMonthPorVivienda(ym, viviendaKey) + managementFeeForMonthPorVivienda(ym, viviendaKey);
      cursorMes += 1;
      if (cursorMes > 12) { cursorMes = 1; cursorAnio += 1; }
    }
    const diferencia = ingresos - gastos;
    const rentabilidad = ingresos > 0 ? (diferencia / ingresos) * 100 : 0;
    return { fechaInicio: fechaInicio, ingresos, gastos, diferencia, rentabilidad };
  }

  function gastoFijoTotal(entry) {
    if (entry == null) return 0;
    if (typeof entry === "object") return Object.values(entry).reduce((s, v) => s + (Number(v) || 0), 0);
    return Number(entry) || 0;
  }
  function fixedExpensesForMonth(ym) {
    const e = expenses[ym] || {};
    const simple = GASTOS_FIJOS_KEYS.reduce((s, k) => s + gastoFijoTotal(e[k]), 0);
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

  function getArrendadorConfig() {
    const s = expenses._settings || {};
    return {
      nombre: s.arrendadorNombre || "",
      documento: s.arrendadorDocumento || "",
      domicilio: s.arrendadorDomicilio || "",
      lugarFirma: s.lugarFirma || "",
    };
  }
  /* La dirección de la vivienda y su referencia catastral también pueden ser distintas por
     vivienda, igual que los gastos. "" representa la única vivienda (o la general, si nunca se
     han usado varias) — así las cuentas con una sola vivienda no ven ningún cambio. */
  function direccionViviendaDesde(settings, viviendaKey) {
    const direcciones = settings.direccionesVivienda || {};
    if (direcciones[viviendaKey]) return direcciones[viviendaKey];
    if (viviendaKey === "") {
      // Compatibilidad con cuentas de antes de que existiera este desglose
      return {
        calle: settings.viviendaCalle || "", numero: settings.viviendaNumero || "",
        piso: settings.viviendaPiso || "", letra: settings.viviendaLetra || "",
        cp: settings.viviendaCP || "", localidad: settings.viviendaLocalidad || "",
        provincia: settings.viviendaProvincia || "", refCatastral: settings.viviendaRefCatastral || "",
      };
    }
    return { calle: "", numero: "", piso: "", letra: "", cp: "", localidad: "", provincia: "", refCatastral: "" };
  }
  function getDireccionVivienda(viviendaKey) {
    return direccionViviendaDesde(expenses._settings || {}, viviendaKey || "");
  }
  function updateDireccionVivienda(viviendaKey, patch) {
    setExpenses(prev => {
      const s = prev._settings || {};
      const actual = direccionViviendaDesde(s, viviendaKey || "");
      const direcciones = { ...(s.direccionesVivienda || {}), [viviendaKey || ""]: { ...actual, ...patch } };
      return { ...prev, _settings: { ...s, direccionesVivienda: direcciones } };
    });
  }
  function updateArrendadorConfig(patch) {
    setExpenses(prev => ({
      ...prev,
      _settings: { ...(prev._settings || {}), ...patch }
    }));
  }
  function descartarOnboarding() {
    setExpenses(prev => ({
      ...prev,
      _settings: { ...(prev._settings || {}), onboardingDescartado: true }
    }));
  }

  async function handleGuardarDatosFiscales(datos) {
    const { error } = await supabase.rpc("actualizar_datos_fiscales", {
      p_nombre_completo: datos.nombreCompleto,
      p_nif: datos.nif,
      p_empresa: datos.empresa,
      p_cif: datos.cif,
      p_direccion: datos.direccion,
      p_telefono: datos.telefono,
      p_acepta_comunicaciones: datos.aceptaComunicaciones,
    });
    if (error) {
      notify("No se pudieron guardar tus datos. Comprueba tu conexión e inténtalo de nuevo.");
      throw error;
    }
    // Estos mismos datos sirven de partida para el contrato — solo faltará añadir
    // la dirección de la vivienda alquilada en cada caso.
    updateArrendadorConfig({
      arrendadorNombre: datos.nombreCompleto,
      arrendadorDocumento: datos.empresa ? datos.cif : datos.nif,
      arrendadorDomicilio: datos.direccion,
    });
    setAccountPlan(prev => ({ ...prev, ...datos }));
    setDatosFiscalesFaltan(false);
    notify("Datos guardados. Ya están listos para usarse en tus contratos.");
  }
  async function fetchPlantillaPersonalizadaBytes(tipo) {
    const prefijo = tipo === "vivienda" ? "plantillaContratoVivienda" : "plantillaContratoHabitacion";
    const path = expenses._settings?.[`${prefijo}Path`];
    if (!path) return null;
    const url = await getDocSignedUrl(path);
    const resp = await fetch(url);
    if (!resp.ok) throw new Error("No se pudo descargar tu plantilla de contrato personalizada.");
    const buf = await resp.arrayBuffer();
    return new Uint8Array(buf);
  }

  function addRoomLabel(name) {
    const clean = String(name || "").trim();
    if (!clean) { notify("Escribe un nombre para la nueva unidad."); return false; }
    if (roomLabels.some(r => r.toLowerCase() === clean.toLowerCase())) {
      notify(`Ya existe una unidad llamada "${clean}".`);
      return false;
    }
    const limite = accountPlan.max_unidades ?? 3;
    const enUso = unidadesPonderadasEnUso(tenants, roomLabels, accountPlan.plan, expenses._settings?.unitTypes);
    if (enUso >= limite) {
      notify(`Tu plan actual permite hasta ${limite} unidades (cada vivienda completa cuenta como varias habitaciones). Mejora de plan para añadir más.`);
      return false;
    }
    setRoomLabels(prev => [...prev, clean]);
    return true;
  }
  function renameRoomLabel(oldName, newName) {
    const clean = String(newName || "").trim();
    if (!clean || clean === oldName) return false;
    if (roomLabels.some(r => r.toLowerCase() === clean.toLowerCase())) {
      notify(`Ya existe una unidad llamada "${clean}".`);
      return false;
    }
    setRoomLabels(prev => prev.map(r => (r === oldName ? clean : r)));
    // Mantener consistentes las fichas (activas e históricas) que usaban el nombre anterior
    setTenants(prev => prev.map(t => (t.habitacion === oldName ? { ...t, habitacion: clean } : t)));
    // El tipo (habitación / vivienda completa) viaja con el nombre nuevo
    setExpenses(prev => {
      const tipos = { ...(prev._settings?.unitTypes || {}) };
      if (tipos[oldName]) {
        tipos[clean] = tipos[oldName];
        delete tipos[oldName];
      }
      return { ...prev, _settings: { ...(prev._settings || {}), unitTypes: tipos } };
    });
    return true;
  }
  function removeRoomLabel(name) {
    const enUso = tenants.some(t => t.activo && matchesRoom(t.habitacion, name));
    if (enUso) {
      notify(`No se puede quitar "${name}" porque tiene un inquilino activo. Libérala primero.`);
      return false;
    }
    setRoomLabels(prev => prev.filter(r => r !== name));
    setExpenses(prev => {
      const tipos = { ...(prev._settings?.unitTypes || {}) };
      delete tipos[name];
      return { ...prev, _settings: { ...(prev._settings || {}), unitTypes: tipos } };
    });
    return true;
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
  function monthIsBeforeGestion(ym) {
    const gestionInicio = toDate(GESTION_INICIO);
    if (!gestionInicio) return false;
    const [y, m] = ym.split("-").map(Number);
    const monthEnd = new Date(y, m, 0); // último día del mes
    return monthEnd < gestionInicio;
  }

  const yearData = useMemo(() => {
    return MESES_CORTOS.map((label, idx) => {
      const ym = ymKey(selYear, idx + 1);
      if (monthIsBeforeGestion(ym)) {
        return { mes: label, ocupacion: 0, ocupacionPct: 0, ingresos: 0, gastosFijos: 0, gestion: 0, gastos: 0, rentabilidad: 0, antesGestion: true };
      }
      return {
        mes: label,
        ocupacion: occupancyForMonth(ym),
        ocupacionPct: Math.round(occupancyPctForMonth(ym) * 10) / 10,
        ingresos: Math.round(incomeForMonth(ym)),
        gastosFijos: Math.round(fixedExpensesForMonth(ym)),
        gestion: Math.round(managementFeeForMonth(ym)),
        gastos: Math.round(totalExpensesForMonth(ym)),
        rentabilidad: Math.round(profitabilityForMonth(ym) * 10) / 10,
        antesGestion: false
      };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tenants, expenses, selYear, roomLabels]);

  const occupancyStats = useMemo(() => {
    const activeMonths = yearData.filter(d => !d.antesGestion);
    const pcts = activeMonths.map(d => d.ocupacionPct);
    const roomMonths = activeMonths.reduce((s, d) => s + d.ocupacion, 0);
    const capacityRoomMonths = roomLabels.length * activeMonths.length;
    return {
      mediaAnual: activeMonths.length ? pcts.reduce((s, p) => s + p, 0) / activeMonths.length : 0,
      roomMonths,
      capacityRoomMonths,
      totalAnual: capacityRoomMonths ? (roomMonths / capacityRoomMonths) * 100 : 0
    };
  }, [yearData, roomLabels]);

  const currentYm = ymKey(now.getFullYear(), now.getMonth() + 1);

  const alerts = useMemo(() => {
    const todayD = toDate(todayISO());
    const impagos = [];
    const finContrato = [];
    const fianzaNoCobrada = [];
    const fianzaNoDevuelta = [];
    const revisionRenta = [];

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
      // Revisión de renta: han pasado 12 meses o más desde la última revisión (o, si nunca se
      // ha revisado, desde el inicio del contrato) — momento habitual para actualizar la renta
      // según el IPC o el IRAV, si el contrato lo prevé.
      if (t.activo) {
        const fechaBase = t.ultimaRevisionRenta || t.fechaInicio;
        const baseD = toDate(fechaBase);
        if (baseD) {
          const meses = (todayD.getFullYear() - baseD.getFullYear()) * 12 + (todayD.getMonth() - baseD.getMonth());
          if (meses >= 12) {
            revisionRenta.push({ ...t, mesesTranscurridos: meses, fechaBase });
          }
        }
      }
    });

    finContrato.sort((a, b) => a.diasRestantes - b.diasRestantes);
    revisionRenta.sort((a, b) => b.mesesTranscurridos - a.mesesTranscurridos);
    return { impagos, finContrato, fianzaNoCobrada, fianzaNoDevuelta, revisionRenta };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tenants]);

  const totalAlerts = alerts.impagos.length + alerts.finContrato.length + alerts.fianzaNoCobrada.length + alerts.fianzaNoDevuelta.length + alerts.revisionRenta.length;
  const incidenciasAbiertasCount = incidencias.filter(i => i.estado !== "cerrada").length;

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

      // Inquilino nuevo: si por lo que sea llega sin historial (no debería pasar, el
      // formulario ya lo prepara), lo iniciamos aquí como red de seguridad.
      if (!Array.isArray(t.historialRenta) || t.historialRenta.length === 0) {
        const mesInicio = t.fechaInicio ? t.fechaInicio.slice(0, 7) : selectedMonth;
        t = { ...t, historialRenta: [{ desde: mesInicio, importe: Number(t.renta) || 0 }] };
      }
      return [...prev, t];
    });
    setShowForm(false);
    setEditingTenant(null);
    notify("Inquilino guardado.");
  }
  function deleteTenant(id) {
    const tenant = tenants.find(p => p.id === id);
    setTenants(prev => prev.filter(p => p.id !== id));
    setConfirmDelete(null);
    if (undoBuffer?.timeoutId) clearTimeout(undoBuffer.timeoutId);
    const timeoutId = setTimeout(() => setUndoBuffer(null), 8000);
    setUndoBuffer({ tenant, timeoutId });
  }
  function undoDeleteTenant() {
    if (!undoBuffer) return;
    clearTimeout(undoBuffer.timeoutId);
    setTenants(prev => [...prev, undoBuffer.tenant]);
    setUndoBuffer(null);
    notify("Eliminación deshecha — el inquilino ha vuelto al histórico.");
  }

  function addIncidencia(datos) {
    const nueva = {
      id: uid(),
      titulo: datos.titulo || "",
      descripcion: datos.descripcion || "",
      habitacion: datos.habitacion || "",
      categoria: datos.categoria || "",
      estado: "abierta",
      fechaCreacion: todayISO(),
      fechaCierre: "",
      presupuestoImporte: datos.presupuestoImporte || "",
      presupuestoDoc: null,
    };
    setIncidencias(prev => [nueva, ...prev]);
    notify("Incidencia registrada.");
    return nueva.id;
  }
  function updateIncidencia(id, patch) {
    setIncidencias(prev => prev.map(inc => {
      if (inc.id !== id) return inc;
      const actualizada = { ...inc, ...patch };
      // Si se marca como cerrada y todavia no tenia fecha de cierre, se rellena sola.
      if (patch.estado === "cerrada" && !inc.fechaCierre) actualizada.fechaCierre = todayISO();
      // Si se reabre, se limpia la fecha de cierre anterior.
      if (patch.estado && patch.estado !== "cerrada") actualizada.fechaCierre = "";
      return actualizada;
    }));
  }
  async function deleteIncidencia(id) {
    const inc = incidencias.find(i => i.id === id);
    if (inc?.presupuestoDoc?.path) {
      try { await deleteDoc(inc.presupuestoDoc.path); } catch (e) { /* no bloqueante */ }
    }
    setIncidencias(prev => prev.filter(i => i.id !== id));
    notify("Incidencia eliminada.");
  }
  async function uploadPresupuestoIncidencia(id, file) {
    try {
      const path = `${accountId}/incidencias/${id}/presupuesto/${Date.now()}_${sanitizeFileName(file.name)}`;
      const anterior = incidencias.find(i => i.id === id)?.presupuestoDoc;
      const meta = await uploadDoc(path, file);
      if (anterior?.path) { try { await deleteDoc(anterior.path); } catch (e) { /* no bloqueante */ } }
      updateIncidencia(id, { presupuestoDoc: meta });
      notify("Presupuesto adjuntado.");
    } catch (e) {
      console.error("Error al subir el presupuesto", e);
      notify("No se pudo subir el presupuesto.");
    }
  }
  async function eliminarPresupuestoIncidencia(id) {
    const inc = incidencias.find(i => i.id === id);
    if (inc?.presupuestoDoc?.path) {
      try { await deleteDoc(inc.presupuestoDoc.path); } catch (e) { /* no bloqueante */ }
    }
    updateIncidencia(id, { presupuestoDoc: null });
  }

  function freeRoom(t, fechaLiberacion) {
    setTenants(prev => prev.map(p => {
      if (p.id !== t.id) return p;
      const patch = { ...p, activo: false };
      // Si la fecha real de salida es anterior a la fecha de fin que tenía el contrato
      // (o si estaba renovado), la ajustamos — si no, el calendario de ocupación seguiría
      // contando la habitación como ocupada hasta la fecha original, aunque ya esté libre.
      const finActual = (p.renovado && p.nuevaFechaFin) ? p.nuevaFechaFin : p.fechaFin;
      if (fechaLiberacion && (!finActual || fechaLiberacion < finActual)) {
        if (p.renovado) patch.nuevaFechaFin = fechaLiberacion;
        else patch.fechaFin = fechaLiberacion;
      }
      return patch;
    }));
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
  function marcarRentaRevisada(tenantId) {
    setTenants(prev => prev.map(t => (t.id === tenantId ? { ...t, ultimaRevisionRenta: todayISO() } : t)));
    notify("Renta marcada como revisada. Próximo aviso en 12 meses.");
  }
  function updateExpenseField(ym, field, value) {
    setExpenses(prev => ({ ...prev, [ym]: { ...(prev[ym] || {}), [field]: value } }));
  }
  function updateExpenseItems(ym, field, items) {
    setExpenses(prev => ({ ...prev, [ym]: { ...(prev[ym] || {}), [field]: items } }));
  }

  async function uploadFacturaGasto(ym, categoria, file) {
    try {
      const path = `${accountId}/gastos/${ym}/${categoria}/${Date.now()}_${sanitizeFileName(file.name)}`;
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

  async function uploadPlantillaContrato(file, tipo) {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      notify("La plantilla debe ser un archivo .docx de Word.");
      return;
    }
    const prefijo = tipo === "vivienda" ? "plantillaContratoVivienda" : "plantillaContratoHabitacion";
    try {
      const path = `${accountId}/config/${prefijo}_${Date.now()}_${sanitizeFileName(file.name)}`;
      const anterior = expenses._settings?.[`${prefijo}Path`];
      await uploadDoc(path, file);
      if (anterior) { try { await deleteDoc(anterior); } catch (e) { /* no bloqueante */ } }
      updateArrendadorConfig({ [`${prefijo}Path`]: path, [`${prefijo}Nombre`]: file.name, [`${prefijo}Fecha`]: new Date().toISOString() });
      notify(tipo === "vivienda" ? "Plantilla de vivienda completa actualizada." : "Plantilla de habitación actualizada.");
    } catch (e) {
      console.error("Error al subir la plantilla", e);
      notify("No se pudo subir la plantilla. Comprueba tu conexión.");
    }
  }

  async function eliminarPlantillaContrato(tipo) {
    const prefijo = tipo === "vivienda" ? "plantillaContratoVivienda" : "plantillaContratoHabitacion";
    const path = expenses._settings?.[`${prefijo}Path`];
    if (!path) return;
    try {
      await deleteDoc(path);
    } catch (e) {
      console.error("Error al borrar la plantilla", e);
    }
    updateArrendadorConfig({ [`${prefijo}Path`]: "", [`${prefijo}Nombre`]: "", [`${prefijo}Fecha`]: "" });
    notify("Plantilla personalizada eliminada. Se usará la plantilla estándar de Susalquia.");
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
            <LayoutGrid size={16} /> {etiquetaUnidades()}
          </button>
          <button className={`rg-nav-item ${view === "inquilinos" ? "active" : ""}`} onClick={() => setView("inquilinos")}>
            <Users size={16} /> Inquilinos
          </button>
          <button className={`rg-nav-item ${view === "gastos" ? "active" : ""}`} onClick={() => setView("gastos")}>
            <Wallet size={16} /> Gastos
          </button>
          <button className={`rg-nav-item ${view === "incidencias" ? "active" : ""}`} onClick={() => setView("incidencias")}>
            <Wrench size={16} /> Incidencias
            {incidenciasAbiertasCount > 0 && <span className="rg-nav-badge">{incidenciasAbiertasCount}</span>}
          </button>
          <button className={`rg-nav-item ${view === "alertas" ? "active" : ""}`} onClick={() => setView("alertas")}>
            <AlertTriangle size={16} /> Alertas
            {totalAlerts > 0 && <span className="rg-nav-badge">{totalAlerts}</span>}
          </button>
          <button className={`rg-nav-item ${view === "historico" ? "active" : ""}`} onClick={() => setView("historico")}>
            <History size={16} /> Histórico
          </button>
          <button className={`rg-nav-item ${view === "micuenta" ? "active" : ""}`} onClick={() => setView("micuenta")}>
            <Settings size={16} /> Mi cuenta
          </button>

          <div className="rg-sidebar-footer">
            <div className="rg-sidebar-footer-info">
              <span className="rg-save-indicator" style={{ color: "var(--accent)" }}><Sparkles size={12} /> Modo demo · datos de ejemplo</span>
              <div style={{ fontSize: 10.5, color: "var(--text-dim)", marginTop: 6, lineHeight: 1.5 }}>
                Nada se guarda de verdad. Prueba a editar, subir documentos o generar un contrato con total libertad.
              </div>
            </div>
            <button className="rg-logout-btn" onClick={reiniciarDemo} title="Reiniciar demo">
              <RotateCcw size={15} /> Reiniciar demo
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
              selYear={selYear} incomeForYear={incomeForYear} deductibleExpensesForYear={deductibleExpensesForYear}
              irpfReduccion={irpfReduccion} setIrpfReduccion={setIrpfReduccion}
              totalHabitaciones={roomLabels.length}
              isGratis={accountPlan.plan === "gratis"}
              tenantsCount={tenants.length}
              arrendadorNombre={getArrendadorConfig().nombre}
              direccionViviendaRellena={direccionViviendaRellena()}
              contratoGeneradoAlgunaVez={!!expenses._settings?.contratoGeneradoAlgunaVez}
              hayGastosIntroducidos={hayGastosIntroducidos()}
              onboardingDescartado={!!expenses._settings?.onboardingDescartado}
              onDescartarOnboarding={descartarOnboarding}
              onGoTo={setView}
              viviendas={listaViviendas()}
              incomeForMonthPorVivienda={incomeForMonthPorVivienda}
              netProfitForMonthPorVivienda={netProfitForMonthPorVivienda}
              profitabilityForMonthPorVivienda={profitabilityForMonthPorVivienda}
              acumuladoDesdeInicioVivienda={acumuladoDesdeInicioVivienda}
              etiquetaUnidades={etiquetaUnidades()}
            />
          )}

          {view === "habitaciones" && (
            <HabitacionesView
              tenants={tenants} selectedMonth={selectedMonth}
              onEdit={openEditTenant} onFree={openConfirmFree} onAddForRoom={openNewTenantForRoom}
              roomLabels={roomLabels} onAddRoom={addRoomLabel} onRenameRoom={renameRoomLabel} onRemoveRoom={removeRoomLabel}
              maxUnidades={accountPlan.max_unidades ?? 3}
              planKey={accountPlan.plan}
              isGratis={accountPlan.plan === "gratis"}
              unitTypes={expenses._settings?.unitTypes}
              onSetTipo={setUnitType}
              etiquetaUnidades={etiquetaUnidades()}
            />
          )}

          {view === "inquilinos" && (
            <InquilinosView
              tenants={tenants} selectedMonth={selectedMonth} shiftMonth={shiftMonth} monthLabel={monthLabel}
              onNew={openNewTenant} onEdit={openEditTenant} onDelete={setConfirmDelete}
              onFree={openConfirmFree} onTogglePago={togglePagoQuick}
              onImportClick={handleImportClick} onExport={handleExport}
              etiquetaUnidades={etiquetaUnidades()}
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
              isGratis={accountPlan.plan === "gratis"}
              tenants={tenants}
              viviendas={listaViviendas()}
              gastoFijoValor={gastoFijoValor}
              updateExpenseFieldVivienda={updateExpenseFieldVivienda}
            />
          )}

          {view === "incidencias" && (
            <IncidenciasView
              incidencias={incidencias} roomLabels={roomLabels}
              onAdd={addIncidencia} onUpdate={updateIncidencia} onDelete={deleteIncidencia}
              onUploadPresupuesto={uploadPresupuestoIncidencia} onEliminarPresupuesto={eliminarPresupuestoIncidencia}
              onViewDoc={viewDoc}
              isGratis={accountPlan.plan === "gratis"}
            />
          )}

          {view === "alertas" && (
            <AlertasView alerts={alerts} onSelectTenant={openEditTenant} isGratis={accountPlan.plan === "gratis"} onMarcarRevisada={marcarRentaRevisada} />
          )}

          {view === "historico" && (
            <HistoricoView tenants={tenants} onEdit={openEditTenant} onExport={handleExportHistorico} onReactivate={reactivateTenant} onDelete={setConfirmDelete} etiquetaUnidades={etiquetaUnidades()} />
          )}

          {view === "micuenta" && (
            <MiCuentaView
              accountPlan={accountPlan} roomLabels={roomLabels} tenants={tenants} session={null} onSignOut={reiniciarDemo}
              onGoToHabitaciones={() => setView("habitaciones")}
              onExportarCopia={exportarCopiaSeguridad}
              onEliminarCuenta={null} eliminandoCuenta={false}
              arrendadorConfig={getArrendadorConfig()} updateArrendadorConfig={updateArrendadorConfig}
              plantillaContratoHabitacionNombre={expenses._settings?.plantillaContratoHabitacionNombre || ""}
              onUploadPlantillaHabitacion={(file) => uploadPlantillaContrato(file, "habitacion")}
              onEliminarPlantillaHabitacion={() => eliminarPlantillaContrato("habitacion")}
              plantillaContratoViviendaNombre={expenses._settings?.plantillaContratoViviendaNombre || ""}
              onUploadPlantillaVivienda={(file) => uploadPlantillaContrato(file, "vivienda")}
              onEliminarPlantillaVivienda={() => eliminarPlantillaContrato("vivienda")}
              viviendas={listaViviendas()}
              getDireccionVivienda={getDireccionVivienda}
              updateDireccionVivienda={updateDireccionVivienda}
              unitTypes={expenses._settings?.unitTypes}
            />
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
          arrendadorConfig={getArrendadorConfig()}
          getDireccionVivienda={getDireccionVivienda}
          onGetPlantillaBytes={fetchPlantillaPersonalizadaBytes}
          getUnitType={getUnitType}
          accountId={accountId}
          isGratis={accountPlan.plan === "gratis"}
          selectedMonth={selectedMonth}
          onContratoGenerado={marcarContratoGenerado}
        />
      )}

      {confirmDelete && (
        <div className="rg-modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div className="rg-modal rg-confirm-box" onClick={(e) => e.stopPropagation()}>
            <div className="rg-modal-title" style={{ marginBottom: 10 }}>¿Eliminar inquilino?</div>
            <p style={{ color: "var(--text-dim)", fontSize: 13.5, marginBottom: 18 }}>
              Se eliminará a <strong style={{ color: "var(--text)" }}>{confirmDelete.nombre} {confirmDelete.apellidos}</strong> y su historial de pagos. Tendrás unos segundos para deshacerlo justo después; pasado ese tiempo, no se podrá recuperar.
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
            <p style={{ color: "var(--text-dim)", fontSize: 13.5, marginBottom: 14 }}>
              <strong style={{ color: "var(--text)" }}>{confirmFree.nombre} {confirmFree.apellidos}</strong> desaparecerá del listado de inquilinos activos
              y pasará al <strong style={{ color: "var(--text)" }}>Histórico</strong>, con todos sus datos guardados. La habitación quedará libre
              para dar de alta a un nuevo inquilino con una ficha en blanco. Puedes deshacerlo desde el Histórico si lo necesitas.
            </p>
            <div className="rg-field" style={{ marginBottom: 4 }}>
              <label className="rg-label">Fecha real de salida</label>
              <input className="rg-input" type="date" value={freeDate} onChange={(e) => setFreeDate(e.target.value)} />
            </div>
            <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginBottom: 18 }}>
              Si el contrato tenía una fecha de fin posterior a esta, se ajustará — así el calendario de
              ocupación y las estadísticas reflejan que la habitación ya está libre a partir de aquí.
            </div>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button className="rg-btn rg-btn-ghost" onClick={() => setConfirmFree(null)}>Cancelar</button>
              <button className="rg-btn" onClick={() => freeRoom(confirmFree, freeDate)}>Liberar habitación</button>
            </div>
          </div>
        </div>
      )}

      {toast && <div className="rg-toast">{toast}</div>}

      {undoBuffer && (
        <div className="rg-toast" style={{ bottom: toast ? 74 : 22, borderColor: "var(--warn)" }}>
          <span>Inquilino "{undoBuffer.tenant.nombre} {undoBuffer.tenant.apellidos}" eliminado.</span>
          <button
            onClick={undoDeleteTenant}
            style={{ background: "none", border: "none", color: "var(--accent)", fontWeight: 700, cursor: "pointer", padding: "2px 4px" }}
          >
            Deshacer
          </button>
        </div>
      )}
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

function YearPicker({ year, onChange }) {
  return (
    <div className="rg-month-picker">
      <button className="rg-icon-btn" onClick={() => onChange(year - 1)} aria-label="Año anterior"><ChevronLeft size={16} /></button>
      <span className="rg-mono">{year}</span>
      <button className="rg-icon-btn" onClick={() => onChange(year + 1)} aria-label="Año siguiente"><ChevronRight size={16} /></button>
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
/* Checklist de bienvenida (onboarding guiado)                          */
/* ------------------------------------------------------------------ */

function OnboardingChecklist({ totalHabitaciones, tenantsCount, arrendadorNombre, direccionViviendaRellena, contratoGeneradoAlgunaVez, hayGastosIntroducidos, isGratis, etiquetaUnidades, onGoTo, onDescartar }) {
  const pasos = [
    {
      hecho: !!arrendadorNombre,
      titulo: "Añade tus datos",
      detalle: "Tu nombre, documento y domicilio — se usan para generar tus contratos.",
      accion: "Ir a Mi cuenta",
      onClick: () => onGoTo("micuenta"),
    },
    {
      hecho: totalHabitaciones > 0,
      titulo: "Crea tu primera unidad",
      detalle: "Ponle el nombre que prefieras, y dinos si es una habitación o una vivienda completa.",
      accion: `Ir a ${etiquetaUnidades}`,
      onClick: () => onGoTo("habitaciones"),
    },
    {
      hecho: direccionViviendaRellena,
      titulo: "Añade datos de la vivienda arrendada",
      detalle: "Dirección y referencia catastral — hacen falta para generar el contrato.",
      accion: "Ir a Mi cuenta",
      onClick: () => onGoTo("micuenta"),
    },
    {
      hecho: tenantsCount > 0,
      titulo: "Crea tu primer inquilino",
      detalle: "Sus datos, su renta y su fianza quedan guardados en su ficha.",
      accion: `Ir a ${etiquetaUnidades}`,
      onClick: () => onGoTo("habitaciones"),
    },
    {
      hecho: contratoGeneradoAlgunaVez,
      titulo: "Genera tu primer contrato",
      detalle: isGratis
        ? "Disponible desde el plan Individual — desde la ficha de cada inquilino."
        : "Desde la ficha del inquilino, con un clic — listo en Word.",
      accion: isGratis ? "Ver planes" : `Ir a ${etiquetaUnidades}`,
      onClick: () => onGoTo(isGratis ? "micuenta" : "habitaciones"),
      bloqueado: isGratis,
    },
    {
      hecho: hayGastosIntroducidos,
      titulo: "Introduce tus primeros gastos",
      detalle: "Luz, agua, IBI, comunidad... para ver tu rentabilidad real cada mes.",
      accion: "Ir a Gastos",
      onClick: () => onGoTo("gastos"),
    },
  ];

  const completados = pasos.filter(p => p.hecho).length;
  const todoListo = completados === pasos.length;

  return (
    <div className="rg-card" style={{ padding: 18, marginBottom: 20, border: "1px solid var(--accent)" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 14.5 }}>
            {todoListo ? "¡Ya tienes lo básico configurado! 🎉" : "Primeros pasos en Susalquia"}
          </div>
          <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 2 }}>
            {completados} de {pasos.length} completados
          </div>
        </div>
        <button
          onClick={onDescartar} title="Ocultar esta guía"
          style={{ background: "none", border: "none", color: "var(--text-dim)", cursor: "pointer", padding: 4 }}
        >
          <X size={16} />
        </button>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 12 }}>
        {pasos.map((paso, i) => (
          <div
            key={i}
            style={{
              display: "flex", alignItems: "center", gap: 12, padding: "10px 12px", borderRadius: 8,
              background: paso.hecho ? "var(--ok-dim)" : "var(--bg)",
              border: `1px solid ${paso.hecho ? "var(--ok)" : "var(--border)"}`
            }}
          >
            <div
              style={{
                width: 22, height: 22, borderRadius: "50%", flexShrink: 0,
                display: "flex", alignItems: "center", justifyContent: "center",
                background: paso.hecho ? "var(--ok)" : "transparent",
                border: paso.hecho ? "none" : "1.5px solid var(--text-dim)"
              }}
            >
              {paso.hecho && <Check size={13} color="#fff" />}
            </div>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 13, fontWeight: 600, textDecoration: paso.hecho ? "line-through" : "none", color: paso.hecho ? "var(--text-dim)" : "var(--text)" }}>
                {paso.titulo}
              </div>
              <div style={{ fontSize: 11.5, color: "var(--text-dim)" }}>{paso.detalle}</div>
            </div>
            {!paso.hecho && (
              <button className="rg-btn rg-btn-ghost" style={{ flexShrink: 0, fontSize: 12, padding: "6px 12px" }} onClick={paso.onClick}>
                {paso.accion}
              </button>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Dashboard                                                            */
/* ------------------------------------------------------------------ */

function DashboardView({ monthLabel, selectedMonth, shiftMonth, incomeForMonth, fixedExpensesForMonth, managementFeeForMonth, totalExpensesForMonth, netProfitForMonth, profitabilityForMonth, yearData, fianzas, occupancyStats, selYear, incomeForYear, deductibleExpensesForYear, irpfReduccion, setIrpfReduccion, totalHabitaciones, isGratis, tenantsCount, arrendadorNombre, direccionViviendaRellena, contratoGeneradoAlgunaVez, hayGastosIntroducidos, onboardingDescartado, onDescartarOnboarding, onGoTo, viviendas, incomeForMonthPorVivienda, netProfitForMonthPorVivienda, profitabilityForMonthPorVivienda, acumuladoDesdeInicioVivienda, etiquetaUnidades }) {
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
  const chartData = yearData.filter(d => !d.antesGestion);

  const tooltipStyle = { background: "#121834", border: "1px solid #262c4d", borderRadius: 8, fontSize: 12, color: "#e7e8f5" };

  return (
    <>
      {!onboardingDescartado && (
        <OnboardingChecklist
          totalHabitaciones={totalHabitaciones}
          tenantsCount={tenantsCount}
          arrendadorNombre={arrendadorNombre}
          direccionViviendaRellena={direccionViviendaRellena}
          contratoGeneradoAlgunaVez={contratoGeneradoAlgunaVez}
          hayGastosIntroducidos={hayGastosIntroducidos}
          isGratis={isGratis}
          etiquetaUnidades={etiquetaUnidades}
          onGoTo={onGoTo}
          onDescartar={onDescartarOnboarding}
        />
      )}

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

      {viviendas.length > 0 && (
        <div style={{ marginBottom: 18 }}>
          <div className="rg-section-title">Rendimiento por vivienda</div>
          <div style={{ display: "grid", gridTemplateColumns: viviendas.length > 1 ? "repeat(2, 1fr)" : "1fr", gap: 14 }}>
            {viviendas.map(v => {
              const ingMes = incomeForMonthPorVivienda(selectedMonth, v);
              const benMes = netProfitForMonthPorVivienda(selectedMonth, v);
              const gastoMes = ingMes - benMes;
              const rentMes = profitabilityForMonthPorVivienda(selectedMonth, v);
              const acum = acumuladoDesdeInicioVivienda(v);

              return (
                <div className="rg-card" style={{ padding: 18 }} key={v}>
                  <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 2 }}>🏠 {v}</div>
                  {acum && (
                    <div style={{ fontSize: 11, color: "var(--text-dim)", marginBottom: 12 }}>
                      Contrato activo desde el {fmtDate(acum.fechaInicio.toISOString ? acum.fechaInicio.toISOString().slice(0, 10) : acum.fechaInicio)}
                    </div>
                  )}

                  <div style={{ fontSize: 11, color: "var(--text-dim)", fontWeight: 700, marginBottom: 6, marginTop: acum ? 0 : 10 }}>
                    {monthLabel.toUpperCase()}
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8, marginBottom: 14 }}>
                    <div>
                      <div style={{ fontSize: 10, color: "var(--text-dim)" }}>Ingresos</div>
                      <div style={{ fontSize: 14, fontWeight: 700, color: "var(--ok)" }}>{fmtMoney(ingMes)}</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 10, color: "var(--text-dim)" }}>Gastos</div>
                      <div style={{ fontSize: 14, fontWeight: 700, color: "var(--danger)" }}>{fmtMoney(gastoMes)}</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 10, color: "var(--text-dim)" }}>Diferencia</div>
                      <div style={{ fontSize: 14, fontWeight: 700, color: benMes >= 0 ? "var(--ok)" : "var(--danger)" }}>{fmtMoney(benMes)}</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 10, color: "var(--text-dim)" }}>Rentabilidad</div>
                      <div style={{ fontSize: 14, fontWeight: 700, color: "var(--accent)" }}>{rentMes.toFixed(1)}%</div>
                    </div>
                  </div>

                  {acum ? (
                    <>
                      <div style={{ fontSize: 11, color: "var(--text-dim)", fontWeight: 700, marginBottom: 6, borderTop: "1px solid var(--border)", paddingTop: 12 }}>
                        ACUMULADO DESDE EL INICIO DEL CONTRATO
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
                        <div>
                          <div style={{ fontSize: 10, color: "var(--text-dim)" }}>Ingresos</div>
                          <div style={{ fontSize: 14, fontWeight: 700, color: "var(--ok)" }}>{fmtMoney(acum.ingresos)}</div>
                        </div>
                        <div>
                          <div style={{ fontSize: 10, color: "var(--text-dim)" }}>Gastos</div>
                          <div style={{ fontSize: 14, fontWeight: 700, color: "var(--danger)" }}>{fmtMoney(acum.gastos)}</div>
                        </div>
                        <div>
                          <div style={{ fontSize: 10, color: "var(--text-dim)" }}>Diferencia</div>
                          <div style={{ fontSize: 14, fontWeight: 700, color: acum.diferencia >= 0 ? "var(--ok)" : "var(--danger)" }}>{fmtMoney(acum.diferencia)}</div>
                        </div>
                        <div>
                          <div style={{ fontSize: 10, color: "var(--text-dim)" }}>Rentabilidad</div>
                          <div style={{ fontSize: 14, fontWeight: 700, color: "var(--accent)" }}>{acum.rentabilidad.toFixed(1)}%</div>
                        </div>
                      </div>
                    </>
                  ) : (
                    <div style={{ fontSize: 11.5, color: "var(--text-dim)", borderTop: "1px solid var(--border)", paddingTop: 12 }}>
                      Sin inquilino activo — el acumulado aparecerá en cuanto la vivienda tenga alguien alojado.
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

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
          <div className="rg-stat-label">Ocupación este mes ({totalHabitaciones} hab.)</div>
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

      {isGratis ? (
        <PlanLockedCard
          titulo="Control de fianzas — disponible desde el plan Individual"
          descripcion="Lleva la cuenta de qué fianzas tienes pendientes de cobrar, cuáles están en tu poder y cuáles ya has devuelto. Mejora tu plan para activarlo."
        />
      ) : (
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
      )}

      <div className="rg-card rg-chart-card">
        <div className="rg-chart-title">Ocupación por mes (%) · {selYear}</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={chartData} margin={{ top: 6, right: 16, left: -18, bottom: 0 }}>
            <CartesianGrid stroke="#262c4d" vertical={false} />
            <XAxis dataKey="mes" tick={{ fill: "#9498b8", fontSize: 11 }} axisLine={{ stroke: "#262c4d" }} tickLine={false} />
            <YAxis domain={[0, 100]} unit="%" tick={{ fill: "#9498b8", fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip contentStyle={tooltipStyle} cursor={{ fill: "#1a2142" }} formatter={(v, n, p) => [`${v}% (${p.payload.ocupacion}/${totalHabitaciones} hab.)`, "Ocupación"]} />
            <Bar dataKey="ocupacionPct" name="Ocupación" fill="#6366f1" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="rg-card rg-chart-card">
        <div className="rg-chart-title">Ingresos vs. gastos · {selYear}</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={chartData} margin={{ top: 6, right: 16, left: -18, bottom: 0 }}>
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
          <LineChart data={chartData} margin={{ top: 6, right: 16, left: -18, bottom: 0 }}>
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
        incomeForYear={incomeForYear}
        deductibleExpensesForYear={deductibleExpensesForYear}
        irpfReduccion={irpfReduccion}
        setIrpfReduccion={setIrpfReduccion}
        isGratis={isGratis}
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

/* Aviso reutilizable para funciones bloqueadas en el plan Gratis */
function PlanLockedCard({ titulo, descripcion, compact }) {
  return (
    <div
      style={{
        padding: compact ? "16px" : "24px 20px", textAlign: "center",
        background: "var(--bg)", border: "1px dashed var(--border)", borderRadius: 10
      }}
    >
      <ShieldAlert size={compact ? 20 : 26} style={{ color: "var(--warn)", marginBottom: 8 }} />
      <div style={{ fontWeight: 600, marginBottom: 4, fontSize: compact ? 13 : 14 }}>{titulo}</div>
      <div style={{ fontSize: 12, color: "var(--text-dim)", maxWidth: 440, margin: "0 auto" }}>{descripcion}</div>
    </div>
  );
}

function IrpfEstimateCard({ selYear, incomeForYear, deductibleExpensesForYear, irpfReduccion, setIrpfReduccion, isGratis }) {
  if (isGratis) {
    return (
      <div className="rg-card rg-chart-card" style={{ paddingBottom: 20 }}>
        <div className="rg-chart-title">Estimación de IRPF · {selYear}</div>
        <div style={{ padding: "16px", textAlign: "center" }}>
          <ShieldAlert size={26} style={{ color: "var(--warn)", marginBottom: 10 }} />
          <div style={{ fontWeight: 600, marginBottom: 6 }}>Disponible desde el plan Individual</div>
          <div style={{ fontSize: 12.5, color: "var(--text-dim)", maxWidth: 420, margin: "0 auto" }}>
            Calcula automáticamente una estimación de tu cuota de IRPF a partir de tus ingresos y gastos
            reales del año, con la reducción legal aplicable. Mejora tu plan para activarlo.
          </div>
        </div>
      </div>
    );
  }

  const ingresos = incomeForYear(selYear);
  const gastos = deductibleExpensesForYear(selYear);
  const rendimientoPrevio = Math.max(0, ingresos - gastos);
  const rendimientoReducido = rendimientoPrevio * (1 - irpfReduccion);
  const cuotaEstimada = calcIrpfProgresivo(rendimientoReducido);
  const tipoMedio = rendimientoReducido > 0 ? (cuotaEstimada / rendimientoReducido) * 100 : 0;

  return (
    <div className="rg-card rg-chart-card" style={{ paddingBottom: 20 }}>
      <div className="rg-chart-title">Estimación de IRPF · {selYear}</div>
      <div style={{ padding: "0 16px" }}>
        <div className="rg-field" style={{ maxWidth: 460 }}>
          <label className="rg-label">Reducción aplicable sobre el rendimiento (Art. 23.2 LIRPF)</label>
          <select className="rg-select" value={irpfReduccion} onChange={(e) => setIrpfReduccion(Number(e.target.value))}>
            {REDUCCIONES_ALQUILER.map(r => <option key={r.value} value={r.value}>{r.label}</option>)}
          </select>
        </div>

        <div className="rg-grid-cards cols-3" style={{ marginTop: 8, marginBottom: 8 }}>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">Ingresos íntegros del año</div>
            <div className="rg-stat-value">{fmtMoney(ingresos)}</div>
          </div>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">Gastos deducibles del año</div>
            <div className="rg-stat-value" style={{ color: "var(--danger)" }}>{fmtMoney(gastos)}</div>
          </div>
          <div className="rg-card rg-stat">
            <div className="rg-stat-label">Rendimiento neto reducido</div>
            <div className="rg-stat-value" style={{ color: "var(--accent)" }}>{fmtMoney(rendimientoReducido)}</div>
          </div>
        </div>
        <div className="rg-grid-cards cols-3" style={{ marginBottom: 8 }}>
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
          se calcula con la escala combinada (estatal + autonómica) de referencia del IRPF 2025/2026, sobre todos los ingresos del año menos
          los gastos deducibles registrados (suministros, comunidad, IBI, reparaciones y gestión), aplicada como si este fuera tu único ingreso
          del año. En la práctica se suma al resto de tus rentas (trabajo, pensión, etc.) y tributa según tu tipo marginal real; además, el tipo
          autonómico exacto depende de tu Comunidad Autónoma. La reducción del alquiler de habitaciones solo aplica si constituye la vivienda
          habitual y permanente del inquilino. Consulta con un asesor fiscal o gestor antes de presentar tu declaración.
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Inquilinos                                                           */
/* ------------------------------------------------------------------ */

function InquilinosView({ tenants, selectedMonth, shiftMonth, monthLabel, onNew, onEdit, onDelete, onFree, onTogglePago, onImportClick, onExport, etiquetaUnidades }) {
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
        💡 Para dar de alta un inquilino nuevo en una unidad concreta, o para liberarla cuando
        alguien se va, te resultará más cómodo usar la sección <strong style={{ color: "var(--text)" }}>{etiquetaUnidades}</strong>.
      </div>

      <div className="rg-card">
        {sorted.length === 0 ? (
          <div className="rg-empty">
            <Users size={30} />
            <div>No hay inquilinos activos ahora mismo.</div>
            <div style={{ fontSize: 12, marginTop: 4 }}>Añade uno nuevo, importa tu Excel, o revisa la sección {etiquetaUnidades}.</div>
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

/* Modal para que cada cliente defina cuántas unidades gestiona y cómo las llama */
function RoomConfigModal({ roomLabels, tenants, onAdd, onRename, onRemove, onClose, maxUnidades, planKey, unitTypes, onSetTipo }) {
  const [nuevoNombre, setNuevoNombre] = useState("");
  const [nuevoTipo, setNuevoTipo] = useState("habitacion");
  const [editando, setEditando] = useState(null); // nombre actual en edición
  const [valorEdicion, setValorEdicion] = useState("");
  const enUsoPonderado = unidadesPonderadasEnUso(tenants, roomLabels, planKey, unitTypes);
  const alLimite = enUsoPonderado >= maxUnidades;
  const hayViviendas = roomLabels.some(label => (unitTypes || {})[label] === "vivienda") || tenants.some(t => t.vivienda && t.vivienda.trim());

  function handleAdd() {
    if (onAdd(nuevoNombre)) {
      onSetTipo(nuevoNombre.trim(), nuevoTipo);
      setNuevoNombre("");
      setNuevoTipo("habitacion");
    }
  }
  function startEdit(label) {
    setEditando(label);
    setValorEdicion(label);
  }
  function confirmEdit() {
    if (onRename(editando, valorEdicion)) setEditando(null);
  }

  return (
    <div className="rg-modal-overlay" onClick={onClose}>
      <div className="rg-modal" style={{ maxWidth: 520 }} onClick={(e) => e.stopPropagation()}>
        <div className="rg-modal-title" style={{ marginBottom: 4 }}>Configurar unidades</div>
        <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 4 }}>
          Añade, renombra o quita las habitaciones o viviendas completas que gestionas — no
          tienen por qué llamarse A1, B2... ponles el nombre que prefieras.
        </div>
        <div style={{ fontSize: 12, color: alLimite ? "var(--warn)" : "var(--text-dim)", marginBottom: 4, fontWeight: 600 }}>
          {Math.round(enUsoPonderado * 10) / 10} de {maxUnidades} unidades usadas en tu plan actual
        </div>
        {hayViviendas && (
          <div style={{ fontSize: 11, color: "var(--text-dim)", marginBottom: 16 }}>
            Cada vivienda completa cuenta como varias habitaciones sueltas, según tu plan.
          </div>
        )}

        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16, maxHeight: 320, overflowY: "auto" }}>
          {roomLabels.map(label => {
            const enUso = tenants.some(t => t.activo && matchesRoom(t.habitacion, label));
            const tipo = (unitTypes || {})[label] || "habitacion";
            return (
              <div key={label} style={{ display: "flex", alignItems: "center", gap: 8, background: "var(--bg)", border: "1px solid var(--border)", borderRadius: 8, padding: "8px 10px" }}>
                {editando === label ? (
                  <>
                    <input
                      className="rg-input" style={{ flex: 1 }} autoFocus value={valorEdicion}
                      onChange={(e) => setValorEdicion(e.target.value)}
                      onKeyDown={(e) => { if (e.key === "Enter") confirmEdit(); if (e.key === "Escape") setEditando(null); }}
                    />
                    <button className="rg-icon-btn" onClick={confirmEdit} title="Guardar"><Check size={14} /></button>
                    <button className="rg-icon-btn" onClick={() => setEditando(null)} title="Cancelar"><X size={14} /></button>
                  </>
                ) : (
                  <>
                    <span style={{ fontSize: 15 }} title={tipo === "vivienda" ? "Vivienda completa" : "Habitación"}>
                      {tipo === "vivienda" ? "🏠" : "🚪"}
                    </span>
                    <span style={{ flex: 1, fontWeight: 600 }}>{label}</span>
                    <select
                      value={tipo} onChange={(e) => onSetTipo(label, e.target.value)}
                      className="rg-select" style={{ fontSize: 11.5, padding: "3px 6px" }}
                      title="Cambiar el tipo de esta unidad"
                    >
                      <option value="habitacion">Habitación</option>
                      <option value="vivienda">Vivienda completa</option>
                    </select>
                    {enUso && <span className="rg-badge rg-badge-ok" style={{ fontSize: 10.5 }}>En uso</span>}
                    <button className="rg-icon-btn" onClick={() => startEdit(label)} title="Renombrar"><Pencil size={13} /></button>
                    <button
                      className="rg-icon-btn" title={enUso ? "Libera la unidad antes de quitarla" : "Quitar"}
                      onClick={() => onRemove(label)} disabled={enUso}
                      style={enUso ? { opacity: .35, cursor: "not-allowed" } : undefined}
                    >
                      <Trash2 size={13} />
                    </button>
                  </>
                )}
              </div>
            );
          })}
        </div>

        <div className="rg-section-title" style={{ fontSize: 12.5, marginBottom: 8 }}>Añadir una unidad nueva</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button
            type="button" onClick={() => setNuevoTipo("habitacion")}
            style={{
              flex: 1, padding: "12px 10px", borderRadius: 10, cursor: "pointer", textAlign: "center",
              border: `2px solid ${nuevoTipo === "habitacion" ? "var(--accent)" : "var(--border)"}`,
              background: nuevoTipo === "habitacion" ? "var(--accent-dim)" : "var(--bg)",
            }}
          >
            <div style={{ fontSize: 22, marginBottom: 4 }}>🚪</div>
            <div style={{ fontWeight: 700, fontSize: 13 }}>Habitación</div>
            <div style={{ fontSize: 10.5, color: "var(--text-dim)", marginTop: 2 }}>Dentro de una vivienda compartida</div>
          </button>
          <button
            type="button" onClick={() => setNuevoTipo("vivienda")}
            style={{
              flex: 1, padding: "12px 10px", borderRadius: 10, cursor: "pointer", textAlign: "center",
              border: `2px solid ${nuevoTipo === "vivienda" ? "var(--accent)" : "var(--border)"}`,
              background: nuevoTipo === "vivienda" ? "var(--accent-dim)" : "var(--bg)",
            }}
          >
            <div style={{ fontSize: 22, marginBottom: 4 }}>🏠</div>
            <div style={{ fontWeight: 700, fontSize: 13 }}>Vivienda completa</div>
            <div style={{ fontSize: 10.5, color: "var(--text-dim)", marginTop: 2 }}>Un piso o casa entera, un solo inquilino</div>
          </button>
        </div>
        <div style={{ fontSize: 11, color: "var(--text-dim)", marginBottom: 10 }}>
          Esta elección decide qué modelo de contrato se genera para esta unidad — puedes
          cambiarla después si te equivocas, con el desplegable de la lista de arriba.
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
          <input
            className="rg-input" placeholder={nuevoTipo === "vivienda" ? "Ej. Piso Alicante, Vivienda Norte..." : "Ej. A1, Ático, Local 3..."} value={nuevoNombre}
            onChange={(e) => setNuevoNombre(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") handleAdd(); }}
            disabled={alLimite}
          />
          <button className="rg-btn" onClick={handleAdd} disabled={alLimite}>
            <Plus size={14} /> Añadir
          </button>
        </div>
        {alLimite && (
          <div style={{ fontSize: 12, color: "var(--warn)", marginBottom: 20 }}>
            Has alcanzado el límite de tu plan. Mejora de plan para añadir más unidades.
          </div>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end" }}>
          <button className="rg-btn" onClick={onClose}>Listo</button>
        </div>
      </div>
    </div>
  );
}

function HabitacionesView({ tenants, selectedMonth, onEdit, onFree, onAddForRoom, roomLabels, onAddRoom, onRenameRoom, onRemoveRoom, maxUnidades, planKey, isGratis, unitTypes, onSetTipo, etiquetaUnidades }) {
  const rooms = roomLabels;
  const ocupadas = rooms.filter(label => tenants.some(t => t.activo && matchesRoom(t.habitacion, label) && overlapsMonth(t, selectedMonth))).length;
  const [calYear, setCalYear] = useState(Number(selectedMonth.split("-")[0]));
  const [showConfig, setShowConfig] = useState(false);

  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">{etiquetaUnidades}</h1>
          <div className="rg-sub">
            {rooms.length === 0
              ? "Todavía no has añadido ninguna habitación o vivienda"
              : `${ocupadas} de ${rooms.length} ocupadas ahora mismo`}
          </div>
        </div>
        <button className="rg-btn rg-btn-ghost" onClick={() => setShowConfig(true)}>
          <Pencil size={13} /> Configurar unidades
        </button>
      </div>

      {showConfig && (
        <RoomConfigModal
          roomLabels={roomLabels} tenants={tenants}
          onAdd={onAddRoom} onRename={onRenameRoom} onRemove={onRemoveRoom}
          onClose={() => setShowConfig(false)}
          maxUnidades={maxUnidades}
          planKey={planKey}
          unitTypes={unitTypes}
          onSetTipo={onSetTipo}
        />
      )}

      {rooms.length === 0 ? (
        <div className="rg-card rg-empty">
          <LayoutGrid size={30} />
          <div style={{ fontWeight: 600, fontSize: 15, marginTop: 10, marginBottom: 4, color: "var(--text)" }}>
            Empieza añadiendo tu primera habitación o vivienda
          </div>
          <div style={{ marginBottom: 16 }}>
            Ponles el nombre que prefieras — no tienen por qué llamarse A1, B2...
          </div>
          <button className="rg-btn" onClick={() => setShowConfig(true)} style={{ margin: "0 auto" }}>
            <Plus size={15} /> Configurar unidades
          </button>
        </div>
      ) : (
      <div className="rg-rooms-grid">
        {rooms.map(label => {
          const tenant = tenants.find(t => t.activo && matchesRoom(t.habitacion, label) && overlapsMonth(t, selectedMonth));
          const paid = tenant && !!(tenant.pagos && tenant.pagos[selectedMonth]);

          // Si no hay nadie viviendo ahí ESTE mes, comprobamos si hay un inquilino ya dado de
          // alta pero cuyo contrato empieza más adelante — así no se ve como "Libre" sin más,
          // y tampoco aparece un pago pendiente que todavía no toca.
          const [selY, selM] = selectedMonth.split("-").map(Number);
          const selMonthEnd = new Date(selY, selM, 0);
          const proximoTenant = !tenant && tenants.find(t => {
            if (!t.activo || !matchesRoom(t.habitacion, label)) return false;
            const inicio = toDate(t.fechaInicio);
            return inicio && inicio > selMonthEnd;
          });

          const esVivienda = (unitTypes || {})[label] === "vivienda";
          const nombreUnidad = esVivienda ? `🏠 ${label}` : `Hab. ${label}`;

          if (!tenant && !proximoTenant) {
            return (
              <div className="rg-room-card rg-room-free" key={label}>
                <div className="rg-room-number">{nombreUnidad}</div>
                <span className="rg-badge rg-badge-neutral" style={{ marginBottom: 10 }}><DoorOpen size={11} /> Libre</span>
                <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 14 }}>Sin inquilino actualmente.</div>
                <button className="rg-btn" style={{ width: "100%", justifyContent: "center" }} onClick={() => onAddForRoom(label)}>
                  <Plus size={15} /> Añadir inquilino
                </button>
              </div>
            );
          }

          if (!tenant && proximoTenant) {
            return (
              <div className="rg-room-card rg-room-free" key={label}>
                <div className="rg-room-number">{nombreUnidad}</div>
                <span className="rg-badge rg-badge-neutral" style={{ marginBottom: 10 }}><DoorOpen size={11} /> Libre este mes</span>
                <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 2 }}>{proximoTenant.nombre} {proximoTenant.apellidos}</div>
                <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 14 }}>
                  Entra el {fmtDate(proximoTenant.fechaInicio)} — todavía no le toca pagar.
                </div>
                <button className="rg-btn rg-btn-ghost" style={{ width: "100%", justifyContent: "center" }} onClick={() => onEdit(proximoTenant)}>
                  <Pencil size={13} /> Ver ficha
                </button>
              </div>
            );
          }

          return (
            <div className="rg-room-card rg-room-occupied" key={label}>
              <div className="rg-room-number">{nombreUnidad}</div>
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
      )}

      <div className="rg-section-title" style={{ marginTop: 28 }}>Información adicional</div>
      <div className="rg-card" style={{ padding: 20 }}>
        <div className="rg-topbar" style={{ marginBottom: 14 }}>
          <div>
            <div className="rg-chart-title" style={{ padding: 0 }}>Calendario anual de ocupación</div>
            <div className="rg-sub" style={{ marginTop: 2 }}>De un vistazo, en qué mes se va quedando libre cada habitación</div>
          </div>
          {!isGratis && <YearPicker year={calYear} onChange={setCalYear} />}
        </div>
        {isGratis ? (
          <PlanLockedCard
            titulo="Disponible desde el plan Individual"
            descripcion="Ve de un vistazo, mes a mes y durante todo el año, en qué momento cada habitación está ocupada o se queda libre. Mejora tu plan para activarlo."
          />
        ) : (
        <>
        <div className="rg-table-wrap">
          <table className="rg-table rg-occupancy-table">
            <thead>
              <tr>
                <th>Hab.</th>
                {MESES_CORTOS.map(m => <th key={m} className="center">{m}</th>)}
              </tr>
            </thead>
            <tbody>
              {rooms.map(label => (
                <tr key={label}>
                  <td className="rg-mono">{label}</td>
                  {MESES_CORTOS.map((_, idx) => {
                    const { percent, tooltip, startsOccupied } = roomMonthOccupancy(tenants, label, calYear, idx + 1);
                    if (percent === null) {
                      return (
                        <td key={idx} title={tooltip}>
                          <div className="rg-occ-track rg-occ-na" />
                        </td>
                      );
                    }
                    return (
                      <td key={idx} title={tooltip}>
                        <div className="rg-occ-track">
                          {(() => {
                            const dispOcupado = percent === 0 ? 0 : Math.max(percent, 6);
                            const dispLibre = 100 - dispOcupado;
                            const ocupadoDiv = dispOcupado > 0 && <div key="o" className="rg-occ-fill rg-occ-fill-ocupado" style={{ width: `${dispOcupado}%` }} />;
                            const libreDiv = dispLibre > 0 && <div key="l" className="rg-occ-fill rg-occ-fill-libre" style={{ width: `${dispLibre}%` }} />;
                            // El tramo que pasa primero dentro del mes se pinta primero (a la izquierda)
                            return startsOccupied ? <>{ocupadoDiv}{libreDiv}</> : <>{libreDiv}{ocupadoDiv}</>;
                          })()}
                        </div>
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 18, marginTop: 12, fontSize: 11.5, color: "var(--text-dim)", flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div className="rg-occ-track" style={{ width: 30, display: "inline-block" }}><div className="rg-occ-fill rg-occ-fill-ocupado" style={{ width: "100%" }} /></div>
            Ocupada
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div className="rg-occ-track" style={{ width: 30, display: "inline-block" }}><div className="rg-occ-fill rg-occ-fill-libre" style={{ width: "100%" }} /></div>
            Libre
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div className="rg-occ-track rg-occ-na" style={{ width: 30, display: "inline-block" }} />
            Antes del inicio de la gestión
          </div>
          <span>· el color se llena en proporción a los días del mes · pasa el cursor sobre un mes para ver la fecha exacta</span>
        </div>
        </>
        )}
      </div>
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Bloque de conceptos itemizados (Reparaciones / Otros gastos)         */
/* ------------------------------------------------------------------ */

function ItemsBlock({ selectedMonth, field, label, conceptos, addLabel, items, factura, onAddItem, onUpdateItem, onRemoveItem, onUploadFactura, onDeleteFactura, onViewDoc, viviendas }) {
  const fileRef = useRef(null);
  const subtotal = items.reduce((s, it) => s + (Number(it.importe) || 0), 0);
  const mostrarVivienda = viviendas && viviendas.length > 1;

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
            {mostrarVivienda && (
              <select
                className="rg-select" value={item.vivienda || ""}
                onChange={(e) => onUpdateItem(field, item.id, { vivienda: e.target.value })}
                title="¿A qué vivienda pertenece este gasto?"
              >
                <option value="">General (todas)</option>
                {viviendas.map(v => <option key={v} value={v}>{v}</option>)}
              </select>
            )}
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

function RepartoSuministrosCard({ tenants, expenses, ym }) {
  const { totalSuministros, totalDias, reparto } = repartoSuministros(tenants, expenses, ym);

  if (totalSuministros === 0) {
    return (
      <div className="rg-card" style={{ padding: 16, marginBottom: 18, textAlign: "center" }}>
        <div style={{ fontSize: 12.5, color: "var(--text-dim)" }}>
          Rellena la luz, el agua o el gas de este mes para ver aquí el reparto entre inquilinos.
        </div>
      </div>
    );
  }

  if (reparto.length === 0) {
    return (
      <div className="rg-card" style={{ padding: 16, marginBottom: 18, textAlign: "center" }}>
        <div style={{ fontSize: 12.5, color: "var(--text-dim)" }}>
          No hay ningún inquilino activo este mes para repartir {fmtMoney(totalSuministros)} de suministros.
        </div>
      </div>
    );
  }

  return (
    <div className="rg-card" style={{ padding: 16, marginBottom: 18 }}>
      <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 12 }}>
        {fmtMoney(totalSuministros)} de luz, agua y gas repartidos entre {reparto.length}{" "}
        {reparto.length === 1 ? "inquilino" : "inquilinos"}, en proporción a los días que ocupó cada uno.
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
        {reparto.map(({ tenant, dias, importe }) => (
          <div
            key={tenant.id}
            style={{
              display: "flex", justifyContent: "space-between", alignItems: "center",
              padding: "8px 12px", background: "var(--bg)", borderRadius: 8, fontSize: 13
            }}
          >
            <div>
              <span style={{ fontWeight: 600 }}>{tenant.nombre} {tenant.apellidos}</span>
              <span style={{ color: "var(--text-dim)", fontSize: 11.5 }}> · Hab. {tenant.habitacion || "—"} · {dias} días</span>
            </div>
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{fmtMoney(importe)}</span>
          </div>
        ))}
      </div>
      <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 10 }}>
        Reparto orientativo para tu propio control — no se cobra ni se descuenta nada automáticamente.
      </div>
    </div>
  );
}

function GastosView({ selectedMonth, monthLabel, shiftMonth, expenses, updateExpenseField, updateExpenseItems, incomeForMonth, fixedExpensesForMonth, managementFeeForMonth, totalExpensesForMonth, netProfitForMonth, yearData, selYear, onUploadFactura, onDeleteFactura, onViewDoc, updateGestionConfig, isGratis, tenants, viviendas, gastoFijoValor, updateExpenseFieldVivienda }) {
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
        {isGratis ? (
          <PlanLockedCard
            titulo="Disponible desde el plan Individual"
            descripcion="Elige si la gestión se cobra como un porcentaje de las rentas o como una cantidad fija al mes. En el plan Gratis se aplica un 15% por defecto. Mejora tu plan para personalizarlo."
            compact
          />
        ) : (
        <>
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
        </>
        )}
      </div>



      <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
        <div className="rg-section-title">Gastos fijos de {monthLabel}</div>
        {viviendas.length > 1 && (
          <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: -6, marginBottom: 10 }}>
            Tienes varias viviendas dadas de alta — indica cuánto corresponde a cada una. Si un
            gasto es compartido entre todas (por ejemplo, un seguro conjunto), ponlo en "General".
          </div>
        )}
        <div className="rg-form-grid">
          {GASTOS_FIJOS_KEYS.map(k => {
            const factura = monthExpenses.facturas?.[k];
            if (viviendas.length > 1) {
              return (
                <div className="rg-field" key={k} style={{ gridColumn: "1 / -1" }}>
                  <label className="rg-label">{GASTOS_FIJOS_LABELS[k]} (€)</label>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                    {["", ...viviendas].map(v => (
                      <div key={v || "general"} style={{ minWidth: 160 }}>
                        <div style={{ fontSize: 10.5, color: "var(--text-dim)", marginBottom: 3 }}>{v || "General"}</div>
                        <input
                          type="number" min="0" step="0.01" className="rg-input"
                          value={gastoFijoValor(monthExpenses[k], v)}
                          onChange={(e) => updateExpenseFieldVivienda(selectedMonth, k, v, e.target.value === "" ? "" : Number(e.target.value))}
                          placeholder="0"
                        />
                      </div>
                    ))}
                  </div>
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
            }
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

        <div className="rg-section-title">Reparto de suministros este mes</div>
        <RepartoSuministrosCard tenants={tenants} expenses={expenses} ym={selectedMonth} />

        <div className="rg-section-title">Reparaciones</div>
        {isGratis ? (
          <PlanLockedCard
            titulo="Disponible desde el plan Individual"
            descripcion="Anota cada reparación por separado (fontanería, electricidad, pintura...) con su importe y su factura adjunta, en vez de un único total. Mejora tu plan para activarlo."
            compact
          />
        ) : (
        <ItemsBlock
          selectedMonth={selectedMonth} field="reparacionesItems" label="Conceptos de reparación"
          conceptos={REPARACIONES_CONCEPTOS} addLabel="Añadir reparación"
          items={monthExpenses.reparacionesItems || []} factura={monthExpenses.facturas?.reparacionesItems}
          onAddItem={addItem} onUpdateItem={updateItem} onRemoveItem={removeItem}
          onUploadFactura={onUploadFactura} onDeleteFactura={onDeleteFactura} onViewDoc={onViewDoc}
          viviendas={viviendas}
        />
        )}

        <div className="rg-section-title">Otros gastos</div>
        {isGratis ? (
          <PlanLockedCard
            titulo="Disponible desde el plan Individual"
            descripcion="Desglosa otros gastos (notaría, impuestos, registro...) uno a uno, con su documento adjunto. Mejora tu plan para activarlo."
            compact
          />
        ) : (
        <ItemsBlock
          selectedMonth={selectedMonth} field="otrosItems" label="Otros conceptos"
          conceptos={OTROS_CONCEPTOS} addLabel="Añadir otro gasto"
          items={monthExpenses.otrosItems || []} factura={monthExpenses.facturas?.otrosItems}
          onAddItem={addItem} onUpdateItem={updateItem} onRemoveItem={removeItem}
          onUploadFactura={onUploadFactura} onDeleteFactura={onDeleteFactura} onViewDoc={onViewDoc}
          viviendas={viviendas}
        />
        )}

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

function AlertasView({ alerts, onSelectTenant, isGratis, onMarcarRevisada }) {
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
        title="Revisión de renta pendiente (IPC / IRAV)" icon={<RefreshCw size={17} />} color="var(--info)"
        items={alerts.revisionRenta} emptyText="Ninguna renta lleva 12 meses o más sin revisar."
        renderItem={(t) => (
          <div className="rg-alert-item" key={t.id}>
            <div onClick={() => onSelectTenant(t)} style={{ cursor: "pointer", flex: 1 }}>
              <div className="rg-alert-item-name">{t.nombre} {t.apellidos}</div>
              <div className="rg-alert-item-sub">
                Habitación {t.habitacion || "—"} · Última revisión: {fmtDate(t.fechaBase)} ({t.mesesTranscurridos} meses)
              </div>
            </div>
            <button
              className="rg-btn rg-btn-ghost" style={{ fontSize: 11.5, padding: "5px 10px", flexShrink: 0 }}
              onClick={(e) => { e.stopPropagation(); onMarcarRevisada(t.id); }}
            >
              Marcar como revisada
            </button>
          </div>
        )}
      />

      {isGratis ? (
        <div className="rg-alert-group">
          <PlanLockedCard
            titulo="Alertas de fianzas — disponible desde el plan Individual"
            descripcion="Avisos automáticos cuando una fianza lleva tiempo sin cobrarse, o cuando toca devolverla al terminar un contrato. Mejora tu plan para activarlo."
          />
        </div>
      ) : (
      <>
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
      )}
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Histórico de inquilinos                                              */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/* Mi cuenta                                                            */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/* Datos fiscales — se piden una vez, al entrar por primera vez         */
/* ------------------------------------------------------------------ */

function DatosFiscalesModal({ onSave, datosIniciales }) {
  const [nombreCompleto, setNombreCompleto] = useState(datosIniciales?.nombre_completo || "");
  const [nif, setNif] = useState(datosIniciales?.nif || "");
  const [tieneEmpresa, setTieneEmpresa] = useState(!!datosIniciales?.empresa);
  const [empresa, setEmpresa] = useState(datosIniciales?.empresa || "");
  const [cif, setCif] = useState(datosIniciales?.cif || "");
  const [direccion, setDireccion] = useState(datosIniciales?.direccion || "");
  const [telefono, setTelefono] = useState(datosIniciales?.telefono || "");
  const [aceptaComunicaciones, setAceptaComunicaciones] = useState(!!datosIniciales?.acepta_comunicaciones);
  const [error, setError] = useState("");
  const [guardando, setGuardando] = useState(false);

  async function handleGuardar() {
    if (!nombreCompleto.trim()) { setError("Escribe tu nombre completo."); return; }
    if (!nif.trim()) { setError("Escribe tu NIF."); return; }
    if (tieneEmpresa && !cif.trim()) { setError("Escribe el CIF de la empresa, o desmarca \"Tengo empresa\" si no aplica."); return; }
    if (!direccion.trim()) { setError("Escribe tu dirección completa."); return; }
    if (!telefono.trim()) { setError("Escribe un teléfono de contacto."); return; }

    setGuardando(true);
    setError("");
    try {
      await onSave({
        nombreCompleto: nombreCompleto.trim(),
        nif: nif.trim(),
        empresa: tieneEmpresa ? empresa.trim() : "",
        cif: tieneEmpresa ? cif.trim() : "",
        direccion: direccion.trim(),
        telefono: telefono.trim(),
        aceptaComunicaciones,
      });
    } catch (e) {
      setError("No se pudieron guardar los datos. Inténtalo de nuevo.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="rg-modal-overlay" style={{ zIndex: 200 }}>
      <div className="rg-modal" style={{ maxWidth: 560 }}>
        <div className="rg-modal-title" style={{ marginBottom: 4 }}>Antes de empezar, completa tus datos</div>
        <p style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 18 }}>
          Los usamos para tu factura y para rellenar automáticamente tus contratos de alquiler —
          solo tendrás que añadir la dirección de cada vivienda. Se piden una sola vez.
        </p>

        <div className="rg-form-grid">
          <div className="rg-field">
            <label className="rg-label">Nombre completo</label>
            <input className="rg-input" value={nombreCompleto} onChange={(e) => setNombreCompleto(e.target.value)} placeholder="Nombre y apellidos" autoFocus />
          </div>
          <div className="rg-field">
            <label className="rg-label">NIF</label>
            <input className="rg-input" value={nif} onChange={(e) => setNif(e.target.value)} placeholder="12345678A" />
          </div>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, marginBottom: 14, cursor: "pointer" }}>
          <input type="checkbox" checked={tieneEmpresa} onChange={(e) => setTieneEmpresa(e.target.checked)} />
          Gestiono los alquileres a través de una empresa
        </label>

        {tieneEmpresa && (
          <div className="rg-form-grid">
            <div className="rg-field">
              <label className="rg-label">Nombre de la empresa</label>
              <input className="rg-input" value={empresa} onChange={(e) => setEmpresa(e.target.value)} placeholder="Razón social" />
            </div>
            <div className="rg-field">
              <label className="rg-label">CIF</label>
              <input className="rg-input" value={cif} onChange={(e) => setCif(e.target.value)} placeholder="B12345678" />
            </div>
          </div>
        )}

        <div className="rg-field">
          <label className="rg-label">Dirección completa</label>
          <input className="rg-input" value={direccion} onChange={(e) => setDireccion(e.target.value)} placeholder="Calle, número, código postal, localidad" />
        </div>

        <div className="rg-field">
          <label className="rg-label">Teléfono de contacto</label>
          <input className="rg-input" type="tel" value={telefono} onChange={(e) => setTelefono(e.target.value)} placeholder="600 000 000" />
        </div>

        <label style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 12, color: "var(--text-dim)", marginBottom: 16, cursor: "pointer", lineHeight: 1.5 }}>
          <input
            type="checkbox" checked={aceptaComunicaciones}
            onChange={(e) => setAceptaComunicaciones(e.target.checked)}
            style={{ marginTop: 2, flexShrink: 0 }}
          />
          Quiero recibir novedades, mejoras de la aplicación y promociones de Susalquia por correo (opcional, puedes cambiarlo cuando quieras desde "Mi cuenta")
        </label>

        {error && (
          <div style={{ background: "var(--danger-dim)", color: "var(--danger)", borderRadius: 8, padding: "9px 12px", fontSize: 12.5, marginBottom: 14 }}>{error}</div>
        )}

        <button className="rg-btn" style={{ width: "100%", justifyContent: "center" }} onClick={handleGuardar} disabled={guardando}>
          {guardando ? "Guardando…" : "Guardar y continuar"}
        </button>
      </div>
    </div>
  );
}

const PLAN_INFO = {
  gratis: { nombre: "Gratis", color: "var(--text-dim)", siguiente: "individual", limiteHabitaciones: 3, limiteViviendas: 1 },
  individual: { nombre: "Individual", color: "var(--accent)", siguiente: "gestor", limiteHabitaciones: 10, limiteViviendas: 2 },
  gestor: { nombre: "Gestor", color: "var(--info)", siguiente: "agencia", limiteHabitaciones: 25, limiteViviendas: 5 },
  agencia: { nombre: "Agencia", color: "var(--ok)", siguiente: null, limiteHabitaciones: null, limiteViviendas: null },
};
/* Una vivienda completa "pesa" más que una habitación suelta, en proporción a lo que cada plan
   ya permite de cada tipo — así si alguien mezcla habitaciones y viviendas en la misma cuenta,
   el límite se reparte de forma justa, sin inventar una proporción que no esté ya reflejada en
   los propios límites del plan. */
function pesoViviendaDelPlan(planKey) {
  const info = PLAN_INFO[planKey] || PLAN_INFO.gratis;
  if (!info.limiteViviendas || !info.limiteHabitaciones) return 5; // Agencia u otros sin límite fijo
  return info.limiteHabitaciones / info.limiteViviendas;
}
/* Cuenta las unidades "ponderadas" en uso: cada habitación sin vivienda asignada vale 1, cada
   vivienda completa distinta vale lo que le corresponda a ese plan. Con cuentas que no usan
   viviendas (la inmensa mayoría), esto da exactamente el mismo número que contar habitaciones
   sueltas, así que no cambia nada para ellas. */
function unidadesPonderadasEnUso(tenants, roomLabels, planKey, unitTypes) {
  const viviendasEnUso = new Set();
  const habitacionesConVivienda = new Set();
  tenants.forEach(t => {
    if (t.vivienda && t.vivienda.trim()) {
      viviendasEnUso.add(t.vivienda.trim());
      if (t.habitacion) habitacionesConVivienda.add(t.habitacion);
    }
  });
  // Las unidades declaradas como "vivienda completa" al crearlas cuentan como su propio grupo,
  // aunque nadie les haya puesto una etiqueta de vivienda a mano.
  roomLabels.forEach(label => {
    if ((unitTypes || {})[label] === "vivienda") {
      viviendasEnUso.add(label);
      habitacionesConVivienda.add(label);
    }
  });
  const habitacionesSueltas = roomLabels.filter(label => !habitacionesConVivienda.has(label)).length;
  return habitacionesSueltas + viviendasEnUso.size * pesoViviendaDelPlan(planKey);
}

function DireccionViviendaFields({ direccion, onChange }) {
  return (
    <div className="rg-form-grid">
      <div className="rg-field">
        <label className="rg-label">Calle</label>
        <input className="rg-input" value={direccion.calle} onChange={(e) => onChange({ calle: e.target.value })} placeholder="Calle Mayor" />
      </div>
      <div className="rg-field">
        <label className="rg-label">Número</label>
        <input className="rg-input" value={direccion.numero} onChange={(e) => onChange({ numero: e.target.value })} placeholder="5" />
      </div>
      <div className="rg-field">
        <label className="rg-label">Piso (opcional)</label>
        <input className="rg-input" value={direccion.piso} onChange={(e) => onChange({ piso: e.target.value })} placeholder="3º" />
      </div>
      <div className="rg-field">
        <label className="rg-label">Letra (opcional)</label>
        <input className="rg-input" value={direccion.letra} onChange={(e) => onChange({ letra: e.target.value })} placeholder="A" />
      </div>
      <div className="rg-field">
        <label className="rg-label">Código postal</label>
        <input className="rg-input" value={direccion.cp} onChange={(e) => onChange({ cp: e.target.value })} placeholder="30001" />
      </div>
      <div className="rg-field">
        <label className="rg-label">Localidad</label>
        <input className="rg-input" value={direccion.localidad} onChange={(e) => onChange({ localidad: e.target.value })} placeholder="Murcia" />
      </div>
      <div className="rg-field">
        <label className="rg-label">Provincia</label>
        <input className="rg-input" value={direccion.provincia} onChange={(e) => onChange({ provincia: e.target.value })} placeholder="Murcia" />
      </div>
      <div className="rg-field">
        <label className="rg-label">Referencia catastral</label>
        <input
          className="rg-input" value={direccion.refCatastral}
          onChange={(e) => onChange({ refCatastral: e.target.value.toUpperCase() })}
          placeholder="1234567AB1234C0001DE" maxLength={20}
          style={{ fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.5px" }}
        />
        <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 4 }}>
          Código de 20 caracteres — lo encuentras en tu último recibo del IBI, o en el
          Catastro (sedecatastro.gob.es).
        </div>
      </div>
    </div>
  );
}

function MiCuentaView({ accountPlan, roomLabels, tenants, session, onSignOut, onGoToHabitaciones, onExportarCopia, onEliminarCuenta, eliminandoCuenta, arrendadorConfig, updateArrendadorConfig, plantillaContratoHabitacionNombre, onUploadPlantillaHabitacion, onEliminarPlantillaHabitacion, plantillaContratoViviendaNombre, onUploadPlantillaVivienda, onEliminarPlantillaVivienda, viviendas, getDireccionVivienda, updateDireccionVivienda, unitTypes }) {
  const planActual = PLAN_INFO[accountPlan.plan] || PLAN_INFO.gratis;
  const siguientePlan = planActual.siguiente ? PLAN_INFO[planActual.siguiente] : null;
  const unidadesUsadas = Math.round(unidadesPonderadasEnUso(tenants, roomLabels, accountPlan.plan, unitTypes) * 10) / 10;
  const limite = accountPlan.max_unidades ?? 3;
  const sobrepasado = unidadesUsadas > limite;
  const alLimite = unidadesUsadas === limite;
  const pct = limite > 0 ? Math.min(100, Math.round((unidadesUsadas / limite) * 100)) : 0;
  const email = session?.user?.email || "";

  const asuntoCorreo = encodeURIComponent(`Quiero mejorar al plan ${siguientePlan?.nombre || ""}`);
  const cuerpoCorreo = encodeURIComponent(
    `Hola,\n\nQuiero mejorar mi cuenta (${email}) del plan ${planActual.nombre} al plan ${siguientePlan?.nombre || ""}.\n\nGracias.`
  );

  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">Mi cuenta</h1>
          <div className="rg-sub">Tu plan, tus datos de acceso y los datos para tus contratos</div>
        </div>
      </div>

      <div className="rg-grid-cards cols-3">
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Plan actual</div>
          <div className="rg-stat-value" style={{ color: planActual.color }}>{planActual.nombre}</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Unidades usadas</div>
          <div className="rg-stat-value" style={{ color: sobrepasado ? "var(--danger)" : "var(--text)" }}>{unidadesUsadas} de {limite}</div>
        </div>
        <div className="rg-card rg-stat">
          <div className="rg-stat-label">Usuarios permitidos</div>
          <div className="rg-stat-value">{accountPlan.max_usuarios ?? 1}</div>
        </div>
      </div>

      {sobrepasado && (
        <div className="rg-card" style={{ padding: 18, marginBottom: 18, border: "1px solid var(--danger)", background: "var(--danger-dim)" }}>
          <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
            <ShieldAlert size={20} style={{ color: "var(--danger)", flexShrink: 0, marginTop: 1 }} />
            <div>
              <div style={{ fontWeight: 700, color: "var(--danger)", marginBottom: 4 }}>
                Tienes más unidades configuradas de las que permite tu plan
              </div>
              <div style={{ fontSize: 12.5, color: "var(--text)", lineHeight: 1.5, marginBottom: 10 }}>
                Tu plan {planActual.nombre} permite hasta {limite}, y ahora mismo tienes {unidadesUsadas} configuradas.
                Esto no bloquea lo que ya tienes creado, pero no podrás añadir ninguna unidad nueva hasta que quites
                alguna de las que sobran o mejores de plan.
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {onGoToHabitaciones && (
                  <button className="rg-btn rg-btn-ghost" onClick={onGoToHabitaciones}>Ir a Configurar unidades</button>
                )}
                {siguientePlan && (
                  <a className="rg-btn" href={`mailto:hola@susalquia.com?subject=${asuntoCorreo}&body=${cuerpoCorreo}`}>
                    Mejorar al plan {siguientePlan.nombre}
                  </a>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
        <div className="rg-section-title">Uso de unidades</div>
        <div style={{ background: "var(--bg)", borderRadius: 999, height: 10, overflow: "hidden", marginBottom: 8, border: "1px solid var(--border)" }}>
          <div
            style={{
              width: `${pct}%`, height: "100%",
              background: sobrepasado || alLimite ? "var(--danger)" : "var(--accent)",
              transition: "width .2s ease"
            }}
          />
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)" }}>
          {unidadesUsadas} de {limite} unidades usadas ({pct}%)
          {sobrepasado ? " — por encima del límite de tu plan, ver aviso arriba" : alLimite ? " — has llegado al límite de tu plan" : ""}
        </div>
      </div>

      <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
        <div className="rg-section-title">Modo demo</div>
        <div style={{ fontSize: 13, color: "var(--text-dim)", marginBottom: 14, maxWidth: 560 }}>
          Estás en una versión de prueba, sin cuenta real ni guardado en la nube. Puedes editar,
          borrar o cambiar lo que quieras — para volver a los datos de ejemplo originales, usa el
          botón de abajo.
        </div>
        <button className="rg-btn rg-btn-ghost" onClick={onSignOut}>
          <RotateCcw size={14} /> Reiniciar demo
        </button>
      </div>

      <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
        <div className="rg-section-title">Datos del arrendador/a y de la vivienda</div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: -6, marginBottom: 12 }}>
          Se usan para rellenar automáticamente el contrato generado en cada ficha de inquilino.
        </div>
        <div className="rg-form-grid">
          <div className="rg-field">
            <label className="rg-label">Nombre completo del arrendador/a</label>
            <input
              className="rg-input" value={arrendadorConfig.nombre}
              onChange={(e) => updateArrendadorConfig({ arrendadorNombre: e.target.value })}
              placeholder="Nombre y apellidos"
            />
          </div>
          <div className="rg-field">
            <label className="rg-label">DNI/NIE del arrendador/a</label>
            <input
              className="rg-input" value={arrendadorConfig.documento}
              onChange={(e) => updateArrendadorConfig({ arrendadorDocumento: e.target.value })}
              placeholder="12345678A"
            />
          </div>
          <div className="rg-field">
            <label className="rg-label">Domicilio del arrendador/a (a efectos de notificaciones)</label>
            <input
              className="rg-input" value={arrendadorConfig.domicilio}
              onChange={(e) => updateArrendadorConfig({ arrendadorDomicilio: e.target.value })}
              placeholder="Calle, número, código postal, localidad"
            />
          </div>
        </div>

        {viviendas.length === 0 ? (
          <>
            <div className="rg-section-title" style={{ marginTop: 20 }}>Dirección de la vivienda que se alquila</div>
            <DireccionViviendaFields
              direccion={getDireccionVivienda("")}
              onChange={(patch) => updateDireccionVivienda("", patch)}
            />
          </>
        ) : (
          <>
            <div className="rg-section-title" style={{ marginTop: 20 }}>Dirección de cada vivienda</div>
            <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: -6, marginBottom: 12 }}>
              Tienes al menos una vivienda completa dada de alta — cada una necesita su propia
              dirección y referencia catastral para que los contratos salgan correctos.
            </div>
            {viviendas.map(v => (
              <div key={v} style={{ marginBottom: 20, paddingBottom: 20, borderBottom: "1px solid var(--border)" }}>
                <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 10, color: "var(--accent)" }}>{v}</div>
                <DireccionViviendaFields
                  direccion={getDireccionVivienda(v)}
                  onChange={(patch) => updateDireccionVivienda(v, patch)}
                />
              </div>
            ))}
            {roomLabels.some(label => (unitTypes || {})[label] !== "vivienda") && (
              <div style={{ marginBottom: 8 }}>
                <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 4 }}>General (habitaciones sueltas)</div>
                <div style={{ fontSize: 11, color: "var(--text-dim)", marginBottom: 10 }}>
                  Usa esta dirección para las habitaciones que no pertenecen a ninguna de las
                  viviendas completas de arriba.
                </div>
                <DireccionViviendaFields
                  direccion={getDireccionVivienda("")}
                  onChange={(patch) => updateDireccionVivienda("", patch)}
                />
              </div>
            )}
          </>
        )}

        <div className="rg-form-grid">
          <div className="rg-field">
            <label className="rg-label">Lugar donde se firma el contrato</label>
            <input
              className="rg-input" value={arrendadorConfig.lugarFirma}
              onChange={(e) => updateArrendadorConfig({ lugarFirma: e.target.value })}
              placeholder="Ej. Murcia"
            />
          </div>
        </div>

        <div className="rg-section-title" style={{ marginTop: 20 }}>Plantillas de contrato</div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 12 }}>
          Susalquia usa un modelo de contrato distinto para habitaciones y para viviendas
          completas, porque legalmente no son el mismo tipo de arrendamiento. Puedes usar las
          plantillas estándar de Susalquia, o subir la tuya propia para cada tipo — debe
          contener las mismas etiquetas (por ejemplo <code className="rg-mono">@@NOMBRE_COMPLETO@@</code>)
          donde quieras que se rellenen los datos automáticamente.
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          <div style={{ border: "1px solid var(--border)", borderRadius: 10, padding: 14 }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 8 }}>🚪 Habitación</div>
            {plantillaContratoHabitacionNombre ? (
              <div className="rg-factura-row">
                <span className="rg-chip"><FileText size={11} /> {plantillaContratoHabitacionNombre}</span>
                <button type="button" className="rg-icon-btn" onClick={onEliminarPlantillaHabitacion} title="Quitar y volver a la plantilla estándar">
                  <Trash2 size={13} />
                </button>
              </div>
            ) : (
              <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 8 }}>
                Usando la plantilla estándar de Susalquia.
              </div>
            )}
            <label className="rg-btn rg-btn-ghost" style={{ display: "inline-flex", marginTop: 8, cursor: "pointer", fontSize: 12.5 }}>
              <Upload size={13} /> {plantillaContratoHabitacionNombre ? "Sustituir" : "Subir la mía (.docx)"}
              <input
                type="file" accept=".docx" style={{ display: "none" }}
                onChange={(e) => { if (e.target.files[0]) onUploadPlantillaHabitacion(e.target.files[0]); e.target.value = ""; }}
              />
            </label>
          </div>

          <div style={{ border: "1px solid var(--border)", borderRadius: 10, padding: 14 }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 8 }}>🏠 Vivienda completa</div>
            {plantillaContratoViviendaNombre ? (
              <div className="rg-factura-row">
                <span className="rg-chip"><FileText size={11} /> {plantillaContratoViviendaNombre}</span>
                <button type="button" className="rg-icon-btn" onClick={onEliminarPlantillaVivienda} title="Quitar y volver a la plantilla estándar">
                  <Trash2 size={13} />
                </button>
              </div>
            ) : (
              <div style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 8 }}>
                Usando la plantilla estándar de Susalquia.
              </div>
            )}
            <label className="rg-btn rg-btn-ghost" style={{ display: "inline-flex", marginTop: 8, cursor: "pointer", fontSize: 12.5 }}>
              <Upload size={13} /> {plantillaContratoViviendaNombre ? "Sustituir" : "Subir la mía (.docx)"}
              <input
                type="file" accept=".docx" style={{ display: "none" }}
                onChange={(e) => { if (e.target.files[0]) onUploadPlantillaVivienda(e.target.files[0]); e.target.value = ""; }}
              />
            </label>
          </div>
        </div>
      </div>

      <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
        <div className="rg-section-title">¿Necesitas ayuda?</div>
        <p style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 14, maxWidth: 560 }}>
          Respuestas a las dudas más habituales: cómo generar un contrato, cómo cambiar de plan,
          cómo exportar tus datos, y más.
        </p>
        <a className="rg-btn rg-btn-ghost" href="https://susalquia.com/ayuda.html" target="_blank" rel="noopener noreferrer">
          Ir al centro de ayuda
        </a>
      </div>

      {siguientePlan && !sobrepasado && (
        <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
          <div className="rg-section-title">¿Necesitas más?</div>
          <p style={{ fontSize: 13, color: "var(--text-dim)", marginBottom: 14, maxWidth: 560 }}>
            El plan <strong style={{ color: "var(--text)" }}>{siguientePlan.nombre}</strong> amplía tu límite hasta{" "}
            {siguientePlan.limiteHabitaciones ? `${siguientePlan.limiteHabitaciones} habitaciones o ${siguientePlan.limiteViviendas} viviendas completas` : "unidades a medida"}
            {planActual.plan === "gratis" && " y desbloquea la generación automática de contratos, la estimación de IRPF, el control de fianzas, el calendario anual de ocupación, la gestión configurable y los gastos desglosados por concepto"}.
          </p>
          <a className="rg-btn" href={`mailto:hola@susalquia.com?subject=${asuntoCorreo}&body=${cuerpoCorreo}`}>
            Mejorar al plan {siguientePlan.nombre}
          </a>
          <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 10 }}>
            De momento el cambio de plan se gestiona por correo — pronto podrás hacerlo tú mismo desde aquí.
          </div>
        </div>
      )}

      <div className="rg-card" style={{ padding: 20, marginBottom: 18 }}>
        <div className="rg-section-title">Tus datos</div>
        <p style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 14, maxWidth: 560 }}>
          Descarga una copia completa de todo lo que tienes guardado en Susalquia: habitaciones,
          inquilinos, gastos y configuración, en un archivo que puedes conservar o llevarte si
          algún día dejas de usar la aplicación. No incluye los documentos adjuntos (contratos,
          DNI, facturas) — esos puedes descargarlos uno a uno desde cada ficha.
        </p>
        <button className="rg-btn rg-btn-ghost" onClick={onExportarCopia}>
          <Download size={14} /> Descargar todos mis datos
        </button>
      </div>

      {onEliminarCuenta && (
        <EliminarCuentaCard onEliminarCuenta={onEliminarCuenta} eliminando={eliminandoCuenta} />
      )}
    </>
  );
}

function EliminarCuentaCard({ onEliminarCuenta, eliminando }) {
  const [mostrarConfirmacion, setMostrarConfirmacion] = useState(false);
  const [texto, setTexto] = useState("");

  return (
    <div className="rg-card" style={{ padding: 20, border: "1px solid var(--danger)" }}>
      <div className="rg-section-title" style={{ color: "var(--danger)" }}>Zona de peligro</div>
      {!mostrarConfirmacion ? (
        <>
          <p style={{ fontSize: 12.5, color: "var(--text-dim)", marginBottom: 14, maxWidth: 560 }}>
            Elimina tu cuenta y todos tus datos de Susalquia: habitaciones, inquilinos, gastos y
            documentos adjuntos. Esta acción no se puede deshacer — descarga antes una copia de
            seguridad si quieres conservar algo.
          </p>
          <button className="rg-btn" style={{ background: "var(--danger)" }} onClick={() => setMostrarConfirmacion(true)}>
            Eliminar mi cuenta
          </button>
        </>
      ) : (
        <>
          <p style={{ fontSize: 12.5, color: "var(--text)", marginBottom: 10, maxWidth: 560 }}>
            Vamos a borrar <strong>todos</strong> tus inquilinos, gastos, habitaciones y documentos.
            Para confirmar, escribe <strong>ELIMINAR</strong> en el campo de abajo.
          </p>
          <input
            className="rg-input" value={texto} onChange={(e) => setTexto(e.target.value)}
            placeholder="ELIMINAR" style={{ maxWidth: 220, marginBottom: 14 }}
          />
          <div style={{ display: "flex", gap: 10 }}>
            <button className="rg-btn rg-btn-ghost" onClick={() => { setMostrarConfirmacion(false); setTexto(""); }}>
              Cancelar
            </button>
            <button
              className="rg-btn" style={{ background: "var(--danger)" }}
              disabled={texto !== "ELIMINAR" || eliminando}
              onClick={onEliminarCuenta}
            >
              {eliminando ? "Eliminando…" : "Confirmar eliminación definitiva"}
            </button>
          </div>
        </>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Gestión de incidencias                                               */
/* ------------------------------------------------------------------ */

function badgeClaseEstado(estado) {
  if (estado === "cerrada") return "rg-badge-ok";
  if (estado === "en_curso") return "rg-badge-warn";
  return "rg-badge-danger";
}

function IncidenciasView({ incidencias, roomLabels, onAdd, onUpdate, onDelete, onUploadPresupuesto, onEliminarPresupuesto, onViewDoc, isGratis }) {
  const [filtro, setFiltro] = useState("todas");
  const [showForm, setShowForm] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const visibles = incidencias.filter(i => filtro === "todas" ? true : i.estado === filtro);
  const contadores = {
    todas: incidencias.length,
    abierta: incidencias.filter(i => i.estado === "abierta").length,
    en_curso: incidencias.filter(i => i.estado === "en_curso").length,
    cerrada: incidencias.filter(i => i.estado === "cerrada").length,
  };

  if (isGratis) {
    return (
      <>
        <div className="rg-topbar">
          <div>
            <h1 className="rg-h1">Incidencias</h1>
            <div className="rg-sub">Averías y reparaciones, con su estado y presupuesto</div>
          </div>
        </div>
        <PlanLockedCard
          titulo="Disponible desde el plan Individual"
          descripcion="Registra averías y reparaciones pendientes, con su estado (abierta, en curso, cerrada) y el presupuesto adjunto de cada una. Mejora tu plan para activarlo."
        />
      </>
    );
  }

  return (
    <>
      <div className="rg-topbar">
        <div>
          <h1 className="rg-h1">Incidencias</h1>
          <div className="rg-sub">Averías y reparaciones, con su estado y presupuesto</div>
        </div>
        <button className="rg-btn" onClick={() => setShowForm(true)}>
          <Plus size={14} /> Nueva incidencia
        </button>
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
        {[
          { key: "todas", label: "Todas" },
          { key: "abierta", label: "Abiertas" },
          { key: "en_curso", label: "En curso" },
          { key: "cerrada", label: "Cerradas" },
        ].map(f => (
          <button
            key={f.key}
            className="rg-btn rg-btn-ghost"
            style={filtro === f.key ? { borderColor: "var(--accent)", color: "var(--accent)" } : undefined}
            onClick={() => setFiltro(f.key)}
          >
            {f.label} ({contadores[f.key]})
          </button>
        ))}
      </div>

      {visibles.length === 0 ? (
        <div className="rg-card" style={{ padding: 24, textAlign: "center" }}>
          <Wrench size={26} style={{ color: "var(--text-dim)", marginBottom: 8 }} />
          <div style={{ fontWeight: 600, marginBottom: 4 }}>
            {incidencias.length === 0 ? "Todavía no has registrado ninguna incidencia" : "No hay incidencias con este filtro"}
          </div>
          <div style={{ fontSize: 12.5, color: "var(--text-dim)" }}>
            Registra aquí averías, reparaciones pendientes o cualquier aviso de una habitación.
          </div>
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {visibles.map(inc => (
            <IncidenciaCard
              key={inc.id} incidencia={inc} roomLabels={roomLabels}
              onUpdate={onUpdate} onDelete={() => setConfirmDelete(inc)}
              onUploadPresupuesto={onUploadPresupuesto} onEliminarPresupuesto={onEliminarPresupuesto}
              onViewDoc={onViewDoc}
            />
          ))}
        </div>
      )}

      {showForm && (
        <IncidenciaFormModal roomLabels={roomLabels} onSave={onAdd} onClose={() => setShowForm(false)} />
      )}

      {confirmDelete && (
        <div className="rg-modal-overlay" onClick={() => setConfirmDelete(null)}>
          <div className="rg-modal" style={{ maxWidth: 420 }} onClick={(e) => e.stopPropagation()}>
            <div className="rg-modal-title" style={{ marginBottom: 10 }}>¿Eliminar esta incidencia?</div>
            <p style={{ color: "var(--text-dim)", fontSize: 13.5, marginBottom: 18 }}>
              Se eliminará <strong style={{ color: "var(--text)" }}>"{confirmDelete.titulo}"</strong> y su presupuesto adjunto, si lo tiene. Esta acción no se puede deshacer.
            </p>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button className="rg-btn rg-btn-ghost" onClick={() => setConfirmDelete(null)}>Cancelar</button>
              <button className="rg-btn rg-btn-danger" onClick={() => { onDelete(confirmDelete.id); setConfirmDelete(null); }}>Eliminar</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function IncidenciaCard({ incidencia, roomLabels, onUpdate, onDelete, onUploadPresupuesto, onEliminarPresupuesto, onViewDoc }) {
  const fileInputRef = useRef(null);
  const estadoInfo = INCIDENCIA_ESTADOS.find(e => e.value === incidencia.estado) || INCIDENCIA_ESTADOS[0];

  return (
    <div className="rg-card" style={{ padding: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, marginBottom: 8 }}>
        <div style={{ flex: 1 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3, flexWrap: "wrap" }}>
            <span style={{ fontWeight: 700, fontSize: 14.5 }}>{incidencia.titulo || "Sin título"}</span>
            <span className={`rg-badge ${badgeClaseEstado(incidencia.estado)}`}>{estadoInfo.label}</span>
          </div>
          <div style={{ fontSize: 12, color: "var(--text-dim)" }}>
            {incidencia.habitacion && <>Habitación {incidencia.habitacion} · </>}
            {incidencia.categoria && <>{incidencia.categoria} · </>}
            Abierta el {fmtDate(incidencia.fechaCreacion)}
            {incidencia.estado === "cerrada" && incidencia.fechaCierre && <> · Cerrada el {fmtDate(incidencia.fechaCierre)}</>}
          </div>
        </div>
        <button className="rg-icon-btn" onClick={onDelete} title="Eliminar" style={{ color: "var(--danger)" }}>
          <Trash2 size={14} />
        </button>
      </div>

      {incidencia.descripcion && (
        <div style={{ fontSize: 13, color: "var(--text)", marginBottom: 12, whiteSpace: "pre-wrap" }}>{incidencia.descripcion}</div>
      )}

      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" }}>
        <div className="rg-field" style={{ marginBottom: 0, minWidth: 160 }}>
          <select
            className="rg-select" value={incidencia.estado}
            onChange={(e) => onUpdate(incidencia.id, { estado: e.target.value })}
          >
            {INCIDENCIA_ESTADOS.map(e => <option key={e.value} value={e.value}>{e.label}</option>)}
          </select>
        </div>

        {incidencia.presupuestoImporte && (
          <span className="rg-badge rg-badge-neutral">Presupuesto: {fmtMoney(incidencia.presupuestoImporte)}</span>
        )}

        {incidencia.presupuestoDoc ? (
          <div className="rg-factura-row">
            <span className="rg-chip" style={{ cursor: "pointer" }} onClick={() => onViewDoc(incidencia.presupuestoDoc.path)}>
              <FileText size={11} /> {incidencia.presupuestoDoc.name}
            </span>
            <button className="rg-icon-btn" onClick={() => onEliminarPresupuesto(incidencia.id)} title="Quitar documento">
              <Trash2 size={12} />
            </button>
          </div>
        ) : (
          <>
            <button className="rg-btn rg-btn-ghost" style={{ fontSize: 12, padding: "6px 12px" }} onClick={() => fileInputRef.current?.click()}>
              <Upload size={12} /> Adjuntar presupuesto
            </button>
            <input
              ref={fileInputRef} type="file" style={{ display: "none" }}
              onChange={(e) => { if (e.target.files[0]) onUploadPresupuesto(incidencia.id, e.target.files[0]); e.target.value = ""; }}
            />
          </>
        )}
      </div>
    </div>
  );
}

function IncidenciaFormModal({ roomLabels, onSave, onClose }) {
  const [titulo, setTitulo] = useState("");
  const [descripcion, setDescripcion] = useState("");
  const [habitacion, setHabitacion] = useState("");
  const [categoria, setCategoria] = useState("");
  const [presupuestoImporte, setPresupuestoImporte] = useState("");
  const [error, setError] = useState("");

  function handleGuardar() {
    if (!titulo.trim()) { setError("Escribe un título breve para la incidencia."); return; }
    onSave({ titulo: titulo.trim(), descripcion: descripcion.trim(), habitacion, categoria, presupuestoImporte });
    onClose();
  }

  return (
    <div className="rg-modal-overlay" onClick={onClose}>
      <div className="rg-modal" style={{ maxWidth: 520 }} onClick={(e) => e.stopPropagation()}>
        <div className="rg-modal-title" style={{ marginBottom: 16 }}>Nueva incidencia</div>

        <div className="rg-field">
          <label className="rg-label">Título</label>
          <input className="rg-input" value={titulo} onChange={(e) => setTitulo(e.target.value)} placeholder="Ej. Grifo de la cocina gotea" autoFocus />
        </div>

        <div className="rg-form-grid">
          <div className="rg-field">
            <label className="rg-label">Habitación o vivienda (opcional)</label>
            <select className="rg-select" value={habitacion} onChange={(e) => setHabitacion(e.target.value)}>
              <option value="">Zona común / sin especificar</option>
              {roomLabels.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
          <div className="rg-field">
            <label className="rg-label">Categoría (opcional)</label>
            <select className="rg-select" value={categoria} onChange={(e) => setCategoria(e.target.value)}>
              <option value="">Sin categoría</option>
              {REPARACIONES_CONCEPTOS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </div>

        <div className="rg-field">
          <label className="rg-label">Descripción (opcional)</label>
          <textarea
            className="rg-input" rows={3} value={descripcion} onChange={(e) => setDescripcion(e.target.value)}
            placeholder="Detalles de la avería o la reparación..." style={{ resize: "vertical" }}
          />
        </div>

        <div className="rg-field">
          <label className="rg-label">Presupuesto estimado en € (opcional)</label>
          <input
            type="number" min="0" step="0.01" className="rg-input" value={presupuestoImporte}
            onChange={(e) => setPresupuestoImporte(e.target.value === "" ? "" : Number(e.target.value))}
            placeholder="0,00"
          />
        </div>

        {error && (
          <div style={{ background: "var(--danger-dim)", color: "var(--danger)", borderRadius: 8, padding: "9px 12px", fontSize: 12.5, marginBottom: 14 }}>{error}</div>
        )}

        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 6 }}>
          <button className="rg-btn rg-btn-ghost" onClick={onClose}>Cancelar</button>
          <button className="rg-btn" onClick={handleGuardar}>Guardar incidencia</button>
        </div>
      </div>
    </div>
  );
}

function HistoricoView({ tenants, onEdit, onExport, onReactivate, onDelete, etiquetaUnidades }) {
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
              Cuando liberes una unidad desde la sección {etiquetaUnidades}, ese inquilino aparecerá aquí.
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
                        <button className="rg-icon-btn" onClick={() => onDelete(t)} title="Eliminar definitivamente" style={{ color: "var(--danger)" }}><Trash2 size={14} /></button>
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

function TenantFormModal({ tenant, onCancel, onSave, notify, onViewDoc, arrendadorConfig, getDireccionVivienda, onGetPlantillaBytes, getUnitType, accountId, isGratis, selectedMonth, onContratoGenerado }) {
  const [form, setForm] = useState(() => {
    let historial = Array.isArray(tenant.historialRenta) ? [...tenant.historialRenta] : [];
    if (historial.length === 0) {
      // Inquilino ya existente antes de que existiera este historial: lo iniciamos con la
      // renta que tiene ahora mismo, vigente desde el inicio de su contrato. Si es un
      // inquilino totalmente nuevo (sin id todavía), esto también le da su primera entrada.
      const mesInicio = tenant.fechaInicio ? tenant.fechaInicio.slice(0, 7) : selectedMonth;
      historial = [{ desde: mesInicio, importe: Number(tenant.renta) || 0 }];
    }
    return { ...tenant, historialRenta: historial };
  });
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
    setForm(prev => {
      const importeNum = Number(value) || 0;
      const historialPrevio = Array.isArray(prev.historialRenta) ? prev.historialRenta : [];
      // El cambio rápido de renta se aplica desde el mes que se esté viendo ahora mismo en la
      // app en adelante — sin tocar ninguna entrada anterior, que conserva la renta que tenía.
      const historial = [
        ...historialPrevio.filter(e => e.desde !== selectedMonth),
        { desde: selectedMonth, importe: importeNum },
      ].sort((a, b) => a.desde.localeCompare(b.desde));
      return {
        ...prev,
        renta: value,
        historialRenta: historial,
        // Mientras la fianza no se haya editado a mano, sigue igualando una mensualidad
        fianzaImporte: fianzaTouched ? prev.fianzaImporte : value
      };
    });
  }
  function setFianza(value) {
    setFianzaTouched(true);
    set("fianzaImporte", value);
  }
  function updateHistorialEntry(idx, field, value) {
    setForm(prev => {
      const historial = [...prev.historialRenta];
      historial[idx] = { ...historial[idx], [field]: field === "importe" ? (Number(value) || 0) : value };
      return { ...prev, historialRenta: historial };
    });
  }
  function removeHistorialEntry(idx) {
    setForm(prev => ({ ...prev, historialRenta: prev.historialRenta.filter((_, i) => i !== idx) }));
  }
  function addHistorialEntry() {
    setForm(prev => ({
      ...prev,
      historialRenta: [...prev.historialRenta, { desde: selectedMonth, importe: Number(prev.renta) || 0 }]
        .sort((a, b) => a.desde.localeCompare(b.desde)),
    }));
  }

  const [generandoContrato, setGenerandoContrato] = useState(false);
  async function handleGenerarContrato() {
    if (isGratis) {
      notify && notify("Generar el contrato en Word está disponible desde el plan Individual. Mejora tu plan para usarlo.");
      return;
    }
    const tipoUnidad = getUnitType ? getUnitType(form.habitacion) : "habitacion";
    const viviendaKey = tipoUnidad === "vivienda" ? (form.vivienda || form.habitacion || "").trim() : (form.vivienda || "").trim();
    const direccionVivienda = getDireccionVivienda(viviendaKey);
    const { obligatorios, opcionales } = validarContrato(form, arrendadorConfig, direccionVivienda);
    if (obligatorios.length) {
      notify && notify(`Faltan datos para generar el contrato: ${obligatorios.join(", ")}.`);
      return;
    }
    if (opcionales.length) {
      const continuar = window.confirm(
        `Faltan estos datos de contacto del inquilino: ${opcionales.join(", ")}.\n\n` +
        `Pulsa "Aceptar" para generar el contrato igualmente, dejando esos datos en blanco para ` +
        `rellenarlos a mano más adelante — o "Cancelar" para completarlos ahora.`
      );
      if (!continuar) return;
    }
    setGenerandoContrato(true);
    try {
      const plantillaBytes = onGetPlantillaBytes ? await onGetPlantillaBytes(tipoUnidad) : null;
      await generarContratoDocx(form, arrendadorConfig, direccionVivienda, plantillaBytes, tipoUnidad);
      notify && notify("Contrato generado. Revisa tu carpeta de descargas.");
      onContratoGenerado && onContratoGenerado();
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
        const path = `${accountId}/tenants/${form.id}/${kind}/${Date.now()}_${sanitizeFileName(file.name)}`;
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
            <label className="rg-label">Unidad (habitación o vivienda)</label>
            <input className="rg-input" value={form.habitacion} onChange={(e) => set("habitacion", e.target.value)} placeholder="Ej. A1, Piso Alicante..." />
            {form.habitacion.trim() && getUnitType && (
              <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 4 }}>
                {getUnitType(form.habitacion.trim()) === "vivienda"
                  ? "🏠 Esta unidad está configurada como vivienda completa."
                  : "🚪 Esta unidad está configurada como habitación."}
              </div>
            )}
          </div>
          <div className="rg-field">
            <label className="rg-label">Agrupar bajo una vivienda (opcional)</label>
            <input
              className="rg-input" value={form.vivienda} onChange={(e) => set("vivienda", e.target.value)}
              placeholder="Ej. Piso Alicante"
              disabled={getUnitType && getUnitType(form.habitacion.trim()) === "vivienda"}
            />
            <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 4 }}>
              {getUnitType && getUnitType(form.habitacion.trim()) === "vivienda"
                ? "No hace falta rellenarlo — esta unidad ya es una vivienda completa por sí sola."
                : "Solo si quieres agrupar varias habitaciones sueltas bajo el nombre de una misma vivienda, para ver su rendimiento conjunto. Si no te aplica, déjalo en blanco."}
            </div>
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
            <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 4 }}>
              Si cambias este importe, se aplicará desde el mes que tengas seleccionado en ese
              momento — los meses ya pasados conservan la renta que tenían entonces.
            </div>
          </div>
        </div>

        <div className="rg-section-title">Historial de renta</div>
        <div style={{ fontSize: 11.5, color: "var(--text-dim)", marginTop: -6, marginBottom: 10 }}>
          Si algún mes pasado no muestra la renta correcta, corrígelo aquí a mano.
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 }}>
          {form.historialRenta.map((entrada, idx) => (
            <div key={idx} style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input
                type="month" className="rg-input" style={{ maxWidth: 150 }}
                value={entrada.desde}
                onChange={(e) => updateHistorialEntry(idx, "desde", e.target.value)}
              />
              <span style={{ color: "var(--text-dim)", fontSize: 12 }}>desde →</span>
              <input
                type="number" min="0" step="0.01" className="rg-input" style={{ maxWidth: 110 }}
                value={entrada.importe}
                onChange={(e) => updateHistorialEntry(idx, "importe", e.target.value)}
              />
              <span style={{ color: "var(--text-dim)", fontSize: 12 }}>€/mes</span>
              {form.historialRenta.length > 1 && (
                <button type="button" className="rg-icon-btn" onClick={() => removeHistorialEntry(idx)} title="Quitar este periodo">
                  <Trash2 size={13} />
                </button>
              )}
            </div>
          ))}
          <button type="button" className="rg-btn rg-btn-ghost" style={{ alignSelf: "flex-start", fontSize: 12, padding: "6px 12px" }} onClick={addHistorialEntry}>
            <Plus size={12} /> Añadir periodo
          </button>
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
              <button
                type="button" className="rg-btn rg-btn-ghost" onClick={handleGenerarContrato} disabled={generandoContrato}
                title={isGratis ? "Disponible desde el plan Individual" : undefined}
              >
                {isGratis ? <ShieldAlert size={13} /> : <FileSignature size={13} />}
                {generandoContrato ? "Generando…" : isGratis ? "Generar contrato 🔒" : "Generar contrato"}
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

/* ------------------------------------------------------------------ */
/* Pantalla de error propia — se muestra si algo se rompe de verdad,    */
/* en vez de dejar una pantalla en blanco                               */
/* ------------------------------------------------------------------ */

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false };
  }
  static getDerivedStateFromError() {
    return { hasError: true };
  }
  componentDidCatch(error, info) {
    console.error("Error no controlado en Susalquia:", error, info);
  }
  handleReload = () => {
    this.setState({ hasError: false });
    window.location.reload();
  };
  render() {
    if (!this.state.hasError) return this.props.children;
    return (
      <div
        style={{
          minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
          background: "#0a0e1f", color: "#e7e8f5", fontFamily: "Inter, -apple-system, sans-serif",
          padding: 24, textAlign: "center"
        }}
      >
        <div style={{ maxWidth: 420 }}>
          <div
            style={{
              width: 52, height: 52, borderRadius: 14, margin: "0 auto 20px auto",
              background: "linear-gradient(135deg, #6366f1, #2f6fed)",
              display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24
            }}
          >
            ⚠️
          </div>
          <div style={{ fontFamily: "Manrope, sans-serif", fontWeight: 800, fontSize: 20, marginBottom: 10 }}>
            Vaya, algo ha ido mal
          </div>
          <div style={{ fontSize: 13.5, color: "#8992b8", lineHeight: 1.6, marginBottom: 26 }}>
            No es culpa de nada que hayas hecho — ha fallado algo inesperado en la aplicación. Tus
            datos están a salvo, guardados en el servidor; prueba a recargar la página. Si el
            problema sigue apareciendo, escríbenos a{" "}
            <a href="mailto:hola@susalquia.com" style={{ color: "#6366f1" }}>hola@susalquia.com</a>{" "}
            y lo revisamos.
          </div>
          <button
            onClick={this.handleReload}
            style={{
              background: "#6366f1", color: "#fff", border: "none", borderRadius: 8,
              padding: "11px 24px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit"
            }}
          >
            Recargar la página
          </button>
        </div>
      </div>
    );
  }
}

export default function App() {
  return (
    <ErrorBoundary>
      <AppInner />
    </ErrorBoundary>
  );
}
