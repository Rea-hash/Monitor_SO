/* Monitor SO — seguimiento de llegada de producto.
   Todo corre en el navegador; los datos se guardan en localStorage. */
'use strict';

// ---------- Configuración ----------
const CFG = {
  // Agencia -> Cedis que la surte. Por defecto CHALCO; excepciones aquí.
  cedisDefault: '15001',
  cedisPorAgencia: { '16006': '15002' },
  cedisNombre: { '15001': 'CHA · 15001 Chalco', '15002': 'TUL · 15002 Tultepec' },
  // Prioridad de asignación por estatus de pago (menor = primero). Dentro: mayor antigüedad primero.
  prioridadPago: {
    'PAGADO': 1, 'FINANCIAMIENTO': 2, 'APARTADO': 3,
    'FACTURACION PENDIENTE': 4, 'PENDIENTE DE REVISION': 5, 'CERRADA': 9
  },
  prioridadOtros: 6,
  // Inventario de agencia que NO cuenta como apto
  garantiaExcluir: ['GARANTIA', 'INFORMATIVO'],
};

const CLS = {
  FI: { k: 'FI', t: 'Facturación inmediata pendiente', b: 'b-fi', kc: 'k-fi' },
  EC: { k: 'EC', t: 'Tránsito – en camino', b: 'b-ec', kc: 'k-ec' },
  PE: { k: 'PE', t: 'Tránsito – por embarcar', b: 'b-pe', kc: 'k-pe' },
  CT: { k: 'CT', t: 'Carga de TO pendiente', b: 'b-ct', kc: 'k-ct' },
  ES: { k: 'ES', t: 'Escalar a Planeación', b: 'b-es', kc: 'k-es' },
};
const CLS_ORDER = ['FI', 'EC', 'PE', 'CT', 'ES'];

const COMENTARIOS = [
  { v: 'RECOGE', t: 'Cliente recoge en fecha compromiso', fecha: true },
  { v: 'PAGO_DOCS', t: 'Cliente pendiente de pago / documentos' },
  { v: 'CAMBIO', t: 'Cliente solicita cambio de modelo/color' },
  { v: 'CANCELO', t: 'Cliente canceló' },
  { v: 'DANADA', t: 'Unidad en piso dañada / en garantía' },
  { v: 'YA_FACT', t: 'Ya facturada, pendiente de reflejarse' },
  { v: 'SIN_RESP', t: 'Gerente no ha respondido' },
  { v: 'OTRO', t: 'Otro', texto: true },
];
const COM_BY = Object.fromEntries(COMENTARIOS.map(c => [c.v, c]));

// ---------- Utilidades ----------
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const norm = v => (v == null ? '' : String(v)).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toUpperCase();
const nkey = v => norm(v).toLowerCase().replace(/\s+/g, ' ');
const code = ub => { const m = String(ub || '').trim().match(/^(\d{4,6})/); return m ? m[1] : norm(ub); };
const num = v => { const n = typeof v === 'number' ? v : parseFloat(String(v ?? '').replace(/,/g, '')); return isFinite(n) ? n : 0; };
const today = () => new Date().toISOString().slice(0, 10);
const daysBetween = (a, b) => Math.round((new Date(b) - new Date(a)) / 864e5);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtD = d => { if (!d) return ''; const x = d instanceof Date ? d : (typeof d === 'number' ? new Date(Math.round((d - 25569) * 864e5)) : new Date(d)); return isNaN(x) ? String(d) : x.toISOString().slice(0, 10); };
function toast(t) { const el = $('#toast'); el.textContent = t; el.classList.add('show'); clearTimeout(toast._t); toast._t = setTimeout(() => el.classList.remove('show'), 2200); }

const LS = {
  get(k, d) { try { const v = localStorage.getItem('monitorso.' + k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem('monitorso.' + k, JSON.stringify(v)); } catch (e) { toast('No se pudo guardar en el navegador: ' + e.message); } },
};

// ---------- Estado ----------
const S = {
  data: LS.get('data', {}),          // {so:{rows,name,at}, to:..., inv:..., ced:...}
  manual: LS.get('manual', []),      // [{to,dest,sos:[],art,qty,added}]
  comments: LS.get('comments', {}),  // {SO: {v, fecha, texto, at}}
  hist: LS.get('hist', {}),          // {SO: {st, since, first}}
  out: LS.get('out', []),            // SOs que salieron
  lastSig: LS.get('lastSig', ''),
  lastSOs: LS.get('lastSOs', []),
  result: [],
  sort: { k: 'prio', dir: 1 },
  f: { cls: new Set(), pay: new Set(), ag: '', q: '', com: '', alert: false },
};

// ---------- Lectura de Excel ----------
const DETECT = {
  so: ['numero de documento', 'estatus pago'],
  to: ['ubicacion destino', 'ubicacion origen'],
  inv: ['numero de serie', 'ubicacion'],
  ced: ['cedis', 'disponible hoy'],
};

function sheetRows(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}

function detect(aoa) {
  for (let i = 0; i < Math.min(aoa.length, 15); i++) {
    const h = (aoa[i] || []).map(nkey);
    for (const [k, req] of Object.entries(DETECT)) {
      if (req.every(r => h.includes(r))) return { k, hi: i, h };
    }
  }
  return null;
}

function objRows(aoa, hi, h) {
  const out = [];
  for (let i = hi + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r) continue;
    const o = {}; let any = false;
    h.forEach((name, j) => { if (name) { o[name] = r[j]; if (r[j] != null && r[j] !== '') any = true; } });
    if (any) out.push(o);
  }
  return out;
}
const isJunk = v => { const s = norm(v); return !s || s === 'TOTAL' || s.startsWith('FILTROS APLICADOS'); };

// Normalización por tipo -> registros compactos (lo que se guarda)
const PARSE = {
  so: rows => rows.filter(r => !isJunk(r['ubicaciones']) && r['numero de documento']).map(r => ({
    ub: String(r['ubicaciones']).trim(), dep: r['departamento'] ?? '', cd: r['cd'] ?? '',
    art: norm(r['articulo']), so: String(r['numero de documento']).trim(), qty: num(r['cantidad']) || 1,
    pago: norm(r['estatus pago']), fecha: fmtD(r['fecha']), ant: num(r['antiguedad']),
    imp: num(r['ingresos totales']), fp: r['forma de pago'] ?? '', estado: r['estado'] ?? '',
    toA: r['to asignada'] ? String(r['to asignada']).trim() : '', toEst: r['estado de to'] ?? '', toF: fmtD(r['fecha creacion de to']),
  })),
  to: rows => rows.filter(r => !isJunk(r['to']) && r['articulo']).map(r => ({
    to: String(r['to']).trim(), fecha: fmtD(r['fecha to']), art: norm(r['articulo']), qty: num(r['cantidad']) || 1,
    orig: r['ubicacion origen'] ?? '', dest: String(r['ubicacion destino'] ?? '').trim(), estado: r['estado'] ?? '',
    if: r['if'] ?? '', fIf: fmtD(r['fecha if']), estIf: r['estado if'] ?? '', sem: r['semaforo'] ?? '', ant: num(r['antiguedad (dias)']),
  })),
  inv: rows => rows.filter(r => !isJunk(r['ubicacion']) && r['articulo']).map(r => ({
    ub: String(r['ubicacion']).trim(), art: norm(r['articulo']), qty: num(r['fisico']) || 1,
    serie: r['numero de serie'] ?? '', sem: r['semaforo'] ?? '', gar: norm(r['tickets.garantia']),
  })),
  ced: rows => rows.filter(r => /^1500[12]\b/.test(String(r['cedis'] || '').trim()) && r['articulo'] && norm(r['articulo']) !== 'TOTAL').map(r => ({
    cedis: String(r['cedis']).trim().slice(0, 5), art: norm(r['articulo']),
    disp: num(r['disponible hoy']), stock: num(r['stock actual']), tin: num(r['transitos entrada']),
  })),
};
const LABEL = { so: 'SOs', to: 'TOs', inv: 'Inventario agencias', ced: 'Inventario Cedis' };

async function readFiles(files) {
  const msgs = [];
  for (const f of files) {
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const aoa = sheetRows(wb);
      const d = detect(aoa);
      if (!d) { msgs.push(`✗ ${f.name}: no reconocí el formato`); continue; }
      const rows = PARSE[d.k](objRows(aoa, d.hi, d.h));
      S.data[d.k] = { rows, name: f.name, at: new Date().toISOString() };
      msgs.push(`✓ ${f.name} → ${LABEL[d.k]} (${rows.length} registros)`);
    } catch (e) { msgs.push(`✗ ${f.name}: ${e.message}`); }
  }
  LS.set('data', S.data);
  purgeManual();
  $('#loadMsg').textContent = msgs.join('  ·  ');
  run(true);
}

// ---------- TOs manuales ----------
function parseManual(txt) {
  const out = [];
  for (const line of txt.split(/\r?\n/)) {
    if (!line.trim()) continue;
    const c = line.split('\t').map(s => s.trim());
    if (c.length < 4 || !/^TO\d+/i.test(c[0])) continue; // salta encabezados
    const sos = (c[2] || '').split(/[,;\s]+/).filter(s => /^SO\d+/i.test(s)).map(s => s.toUpperCase());
    out.push({ to: c[0].toUpperCase(), dest: c[1], sos, art: norm(c[3]), qty: num(c[4]) || 1, added: today() });
  }
  return out;
}
function purgeManual() {
  const inFile = new Set((S.data.to?.rows || []).map(r => r.to.toUpperCase()));
  const before = S.manual.length;
  S.manual = S.manual.filter(m => !inFile.has(m.to));
  if (S.manual.length !== before) toast(`${before - S.manual.length} línea(s) de TO manual ya aparecen en el Excel y se quitaron`);
  LS.set('manual', S.manual);
}

// ---------- Motor de clasificación ----------
function prio(r) { return CFG.prioridadPago[r.pago] ?? CFG.prioridadOtros; }
const cedisDe = ub => CFG.cedisPorAgencia[code(ub)] || CFG.cedisDefault;

function classify() {
  const sos = (S.data.so?.rows || []).slice();
  const inv = S.data.inv?.rows || [];
  const tos = S.data.to?.rows || [];
  const ced = S.data.ced?.rows || [];

  // Orden de prioridad
  sos.sort((a, b) => prio(a) - prio(b) || b.ant - a.ant || (a.fecha || '').localeCompare(b.fecha || '') || a.so.localeCompare(b.so));

  // Pools de inventario de agencia
  const invApt = {}, invNoApt = {};
  for (const u of inv) {
    const k = code(u.ub) + '|' + u.art;
    const noApto = CFG.garantiaExcluir.some(g => u.gar.startsWith(g));
    if (noApto) invNoApt[k] = (invNoApt[k] || 0) + u.qty; else invApt[k] = (invApt[k] || 0) + u.qty;
  }

  // Líneas de tránsito (Excel + manuales)
  const lines = [];
  for (const t of tos) {
    const enCamino = norm(t.estado).startsWith('RECEPCION') || norm(t.estIf) === 'ENVIADO';
    lines.push({ ...t, ag: code(t.dest), left: t.qty, stage: enCamino ? 'EC' : 'PE', src: 'Excel', sos: [] });
  }
  for (const m of S.manual) {
    lines.push({ to: m.to, art: m.art, qty: m.qty, dest: m.dest, ag: code(m.dest), left: m.qty, stage: 'PE', src: 'Manual', sos: m.sos, estado: 'TO manual (correo)' });
  }
  const toIdx = {};
  for (const l of lines) (toIdx[l.to.toUpperCase()] ||= []).push(l);

  // Cedis disponible
  const cedPool = {}, cedIn = {};
  for (const c of ced) {
    const k = c.cedis + '|' + c.art;
    cedPool[k] = (cedPool[k] || 0) + Math.max(0, c.disp);
    cedIn[k] = (cedIn[k] || 0) + Math.max(0, c.tin);
  }

  const res = sos.map((r, i) => ({ ...r, prio: i + 1, ag: code(r.ub), cls: null, det: '', alerts: [], cedis: cedisDe(r.ub) }));
  // Alertas de datos que no cuadran. Departamento "normal" de cada agencia = el más frecuente.
  const depMode = {};
  for (const r of res) { const m = (depMode[r.ag] ||= {}); m[r.dep] = (m[r.dep] || 0) + 1; }
  for (const ag in depMode) depMode[ag] = Object.entries(depMode[ag]).sort((a, b) => b[1] - a[1])[0][0];
  for (const r of res) {
    if (r.dep && r.dep !== depMode[r.ag]) r.alerts.push(`Departamento distinto: ${r.dep} (la agencia usa ${depMode[r.ag]})`);
    const cdEsp = r.cedis === '15002' ? 'TUL' : 'CHL';
    if (r.cd && norm(r.cd) !== cdEsp) r.alerts.push(`CD en SO = ${r.cd}, se esperaba ${cdEsp}`);
  }

  // Pasada A: inventario en agencia
  for (const r of res) {
    const k = r.ag + '|' + r.art;
    if ((invApt[k] || 0) >= r.qty) {
      invApt[k] -= r.qty; r.cls = 'FI'; r.det = 'Unidad en piso de la agencia';
    }
  }

  // Pasada B: TOs vinculadas directamente a la SO (TO Asignada o Nota de TO manual)
  const linkLine = (r, l) => { l.left -= r.qty; r.cls = l.stage; r.det = descLine(l, true); };
  for (const r of res) {
    if (r.cls) continue;
    // 1) TO Asignada de la SO
    if (r.toA) {
      const ls = toIdx[r.toA.toUpperCase()];
      if (ls) {
        const l = ls.find(x => x.art === r.art && x.ag === r.ag && x.left >= r.qty);
        if (l) { linkLine(r, l); continue; }
        const otros = [...new Set(ls.map(x => x.art))].join(', ');
        if (!ls.some(x => x.art === r.art)) r.alerts.push(`${r.toA} trae otro artículo: ${otros}`);
      } else if (norm(r.toEst) === 'RECIBIDA') {
        r.alerts.push(`${r.toA} marcada Recibida pero no hay unidad apta en inventario`);
      } else {
        r.alerts.push(`${r.toA} (${r.toEst || 'sin estado'}) no está en el archivo de TOs`);
      }
    }
    // 2) TO manual cuya Nota menciona la SO
    const l = lines.find(x => x.sos.includes(r.so.toUpperCase()) && x.art === r.art && x.ag === r.ag && x.left >= r.qty);
    if (l) linkLine(r, l);
  }

  // Pasada C: tránsito general -> cedis -> escalar
  for (const r of res) {
    if (r.cls) continue;
    const cand = lines.filter(x => x.ag === r.ag && x.art === r.art && x.left >= r.qty)
      .sort((a, b) => (a.stage === 'EC' ? 0 : 1) - (b.stage === 'EC' ? 0 : 1) || (a.fecha || '').localeCompare(b.fecha || ''));
    if (cand.length) { const l = cand[0]; l.left -= r.qty; r.cls = l.stage; r.det = descLine(l, false); continue; }
    const ck = r.cedis + '|' + r.art;
    if ((cedPool[ck] || 0) >= r.qty) {
      cedPool[ck] -= r.qty; r.cls = 'CT'; r.det = `Disponible en ${CFG.cedisNombre[r.cedis]} (quedan ${cedPool[ck]})`;
      continue;
    }
    r.cls = 'ES';
    if ((cedIn[ck] || 0) >= r.qty) {
      cedIn[ck] -= r.qty;
      r.det = `Sin disponible hoy en ${CFG.cedisNombre[r.cedis]} · cubierta por tránsito de entrada al cedis (quedan ${cedIn[ck]})`;
    } else {
      r.det = `Sin inventario en agencia, tránsito ni ${CFG.cedisNombre[r.cedis]} · sin tránsito de entrada al cedis`;
    }
  }

  // Nota de inventario no apto
  for (const r of res) {
    const k = r.ag + '|' + r.art;
    if (r.cls !== 'FI' && invNoApt[k]) r.alerts.push(`Sin inventario apto: ${invNoApt[k]} unidad(es) en piso con garantía/informativo`);
  }
  return res;
}
function descLine(l, vinculada) {
  const p = [l.to];
  if (l.src === 'Manual') p.push('manual');
  if (l.if) p.push(`${l.if}${l.fIf ? ' ' + l.fIf : ''}`);
  if (l.fecha) p.push(`TO ${l.fecha}`);
  return (vinculada ? 'Vinculada · ' : '') + p.join(' · ');
}

// ---------- Historial / salidas ----------
function updateHistory(res, fresh) {
  const t = today();
  for (const r of res) {
    const h = S.hist[r.so];
    if (!h) S.hist[r.so] = { st: r.cls, since: t, first: t };
    else if (h.st !== r.cls) { h.st = r.cls; h.since = t; }
  }
  const sig = res.map(r => r.so).sort().join(',');
  if (fresh && S.data.so && sig !== S.lastSig) {
    const cur = new Set(res.map(r => r.so));
    if (S.lastSOs.length) {
      S.out = S.lastSOs.filter(x => !cur.has(x.so)).map(x => ({ ...x, salio: t, com: S.comments[x.so] || null }));
    }
    S.lastSOs = res.map(r => ({ so: r.so, ub: r.ub, art: r.art, pago: r.pago, cls: r.cls, fecha: r.fecha }));
    S.lastSig = sig;
    LS.set('out', S.out); LS.set('lastSOs', S.lastSOs); LS.set('lastSig', S.lastSig);
  }
  LS.set('hist', S.hist);
}

// ---------- Render ----------
function run(fresh) {
  renderSlots();
  if (!S.data.so) { S.result = []; renderAll(); return; }
  S.result = classify();
  updateHistory(S.result, fresh);
  renderAll();
}

function renderSlots() {
  for (const el of $$('.slot')) {
    const d = S.data[el.dataset.k];
    el.classList.toggle('ok', !!d);
    el.querySelector('.st').textContent = d ? `${d.rows.length} reg · ${d.name} · ${new Date(d.at).toLocaleString('es-MX', { dateStyle: 'short', timeStyle: 'short' })}` : 'Sin cargar';
  }
}

function filtered() {
  const f = S.f, q = norm(f.q);
  return S.result.filter(r => {
    if (f.cls.size && !f.cls.has(r.cls)) return false;
    if (f.pay.size && !f.pay.has(r.pago)) return false;
    if (f.ag && r.ag !== f.ag) return false;
    if (f.alert && !r.alerts.length) return false;
    const c = S.comments[r.so];
    if (f.com === '__none' && c?.v) return false;
    if (f.com === '__any' && !c?.v) return false;
    if (f.com && !f.com.startsWith('__') && c?.v !== f.com) return false;
    if (q && !norm([r.so, r.art, r.ub, r.det, r.toA, r.dep].join(' ')).includes(q)) return false;
    return true;
  });
}

function renderAll() {
  renderKpis(); renderFilters(); renderTable(); renderAg(); renderMan(); renderOut();
}

function renderKpis() {
  const cnt = Object.fromEntries(CLS_ORDER.map(k => [k, 0]));
  S.result.forEach(r => cnt[r.cls]++);
  const all = `<div class="kpi k-all ${S.f.cls.size ? '' : 'sel'}" data-k=""><div class="n">${S.result.length}</div><div class="t">SOs totales</div></div>`;
  $('#kpis').innerHTML = all + CLS_ORDER.map(k => `<div class="kpi ${CLS[k].kc} ${S.f.cls.has(k) && S.f.cls.size === 1 ? 'sel' : ''}" data-k="${k}"><div class="n">${cnt[k]}</div><div class="t">${CLS[k].t}</div></div>`).join('');
}

function renderFilters() {
  // agencias
  const ags = [...new Map(S.result.map(r => [r.ag, r.ub])).entries()].sort((a, b) => a[1].localeCompare(b[1]));
  const sel = $('#fAg'), cur = S.f.ag;
  sel.innerHTML = '<option value="">Todas las agencias</option>' + ags.map(([c, n]) => `<option value="${c}">${esc(n)}</option>`).join('');
  sel.value = cur;
  // comentarios
  const fc = $('#fCom'), curc = S.f.com;
  fc.innerHTML = '<option value="">Comentario: todos</option><option value="__none">Sin comentario</option><option value="__any">Con comentario</option>' + COMENTARIOS.map(c => `<option value="${c.v}">${esc(c.t)}</option>`).join('');
  fc.value = curc;
  // chips clasificación
  const base = S.result;
  $('#fCls').innerHTML = CLS_ORDER.map(k => `<button class="chip ${S.f.cls.has(k) ? 'on' : ''}" data-cls="${k}">${CLS[k].t}<span class="c">${base.filter(r => r.cls === k).length}</span></button>`).join('');
  // chips estatus pago (dinámico)
  const pays = [...new Set(base.map(r => r.pago))].sort((a, b) => (CFG.prioridadPago[a] ?? 6) - (CFG.prioridadPago[b] ?? 6) || a.localeCompare(b));
  $('#fPay').innerHTML = pays.map(p => `<button class="chip ${S.f.pay.has(p) ? 'on' : ''}" data-pay="${esc(p)}">${esc(p || '(vacío)')}<span class="c">${base.filter(r => r.pago === p).length}</span></button>`).join('');
}

const COLS = [
  { k: 'prio', t: '#', num: true },
  { k: 'ub', t: 'Agencia' },
  { k: 'so', t: 'SO' },
  { k: 'art', t: 'Artículo' },
  { k: 'pago', t: 'Estatus pago' },
  { k: 'ant', t: 'Antig.', num: true },
  { k: 'cls', t: 'Clasificación' },
  { k: 'det', t: 'Detalle / alertas' },
  { k: 'dias', t: 'Días en estatus', num: true },
  { k: 'com', t: 'Comentario del gerente' },
];
function sortVal(r, k) {
  if (k === 'cls') return CLS_ORDER.indexOf(r.cls);
  if (k === 'dias') return diasEst(r);
  if (k === 'com') return S.comments[r.so]?.v || '~';
  return r[k] ?? '';
}
const diasEst = r => { const h = S.hist[r.so]; return h ? daysBetween(h.since, today()) : 0; };

function renderTable() {
  const tb = $('#tblSO');
  tb.querySelector('thead').innerHTML = '<tr>' + COLS.map(c => `<th class="${c.num ? 'num' : ''}" data-sort="${c.k}">${c.t}${S.sort.k === c.k ? (S.sort.dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('') + '</tr>';
  const rows = filtered();
  const { k, dir } = S.sort;
  rows.sort((a, b) => { const x = sortVal(a, k), y = sortVal(b, k); return (x > y ? 1 : x < y ? -1 : 0) * dir || a.prio - b.prio; });
  if (!S.data.so) {
    tb.querySelector('tbody').innerHTML = `<tr><td colspan="${COLS.length}" class="det">Carga los archivos para comenzar. Mínimo el de SOs; para clasificar bien se necesitan los 4.</td></tr>`;
    $('#soFoot').textContent = ''; return;
  }
  tb.querySelector('tbody').innerHTML = rows.map(r => {
    const c = S.comments[r.so] || {};
    const def = COM_BY[c.v];
    const h = S.hist[r.so];
    return `<tr data-so="${esc(r.so)}">
      <td class="num">${r.prio}</td>
      <td>${esc(r.ub)}</td>
      <td class="mono">${esc(r.so)}${r.qty > 1 ? ` ×${r.qty}` : ''}</td>
      <td class="mono">${esc(r.art)}</td>
      <td>${esc(r.pago)}<div class="det">${esc(r.fecha)}</div></td>
      <td class="num">${r.ant}</td>
      <td><span class="badge ${CLS[r.cls].b}">${CLS[r.cls].t}</span></td>
      <td><span class="det">${esc(r.det)}</span>${r.alerts.map(a => `<span class="alert">⚠ ${esc(a)}</span>`).join('')}</td>
      <td class="num" title="${h ? 'Desde ' + h.since + ' · visto por 1ª vez ' + h.first : ''}">${diasEst(r)}</td>
      <td><div class="com">
        <select data-com="v"><option value="">—</option>${COMENTARIOS.map(o => `<option value="${o.v}" ${c.v === o.v ? 'selected' : ''}>${esc(o.t)}</option>`).join('')}</select>
        ${def?.fecha ? `<input type="date" data-com="fecha" value="${esc(c.fecha || '')}">` : ''}
        ${c.v ? `<input type="text" data-com="texto" placeholder="${def?.texto ? 'Describe…' : 'Nota opcional'}" value="${esc(c.texto || '')}" maxlength="120">` : ''}
        ${c.at ? `<span class="meta">Actualizado ${esc(c.at)}</span>` : ''}
      </div></td></tr>`;
  }).join('') || `<tr><td colspan="${COLS.length}" class="det">Sin resultados con estos filtros.</td></tr>`;
  $('#soFoot').textContent = `${rows.length} de ${S.result.length} SOs · Prioridad: Pagado → Financiamiento → Apartado → Facturación pendiente → Pendiente de revisión → otros; dentro de cada grupo, mayor antigüedad primero.`;
}

function renderAg() {
  const m = {};
  for (const r of S.result) {
    const a = (m[r.ag] ||= { ub: r.ub, cedis: r.cedis, c: {}, tot: 0, al: 0 });
    a.c[r.cls] = (a.c[r.cls] || 0) + 1; a.tot++; if (r.alerts.length) a.al++;
  }
  const ags = Object.entries(m).sort((a, b) => a[1].ub.localeCompare(b[1].ub));
  const tot = Object.fromEntries(CLS_ORDER.map(k => [k, 0]));
  const cell = (ag, k, n) => n ? `<a href="#" data-ag="${ag}" data-cls="${k}">${n}</a>` : '<span class="zero">0</span>';
  $('#tblAg').innerHTML = `<thead><tr><th>Agencia</th><th>Cedis</th>${CLS_ORDER.map(k => `<th class="num"><span class="badge ${CLS[k].b}">${CLS[k].t}</span></th>`).join('')}<th class="num">Total</th><th class="num">Con alertas</th></tr></thead><tbody>` +
    ags.map(([ag, a]) => { CLS_ORDER.forEach(k => tot[k] += a.c[k] || 0); return `<tr class="agn"><td>${esc(a.ub)}</td><td>${a.cedis === '15002' ? 'TUL' : 'CHA'}</td>${CLS_ORDER.map(k => `<td class="num">${cell(ag, k, a.c[k] || 0)}</td>`).join('')}<td class="num">${cell(ag, '', a.tot)}</td><td class="num">${a.al || '<span class="zero">0</span>'}</td></tr>`; }).join('') +
    `<tr><td><b>Total</b></td><td></td>${CLS_ORDER.map(k => `<td class="num"><b>${tot[k]}</b></td>`).join('')}<td class="num"><b>${S.result.length}</b></td><td></td></tr></tbody>`;
}

function renderMan() {
  $('#cntMan').textContent = S.manual.length || '';
  const regs = new Set(S.result.map(r => r.ag));
  $('#tblMan').innerHTML = S.manual.length ? `<thead><tr><th>TO</th><th>Destino</th><th>SOs (nota)</th><th>Artículo</th><th class="num">Cant.</th><th>Región</th><th>Agregada</th><th></th></tr></thead><tbody>` +
    S.manual.map((m, i) => `<tr><td class="mono">${esc(m.to)}</td><td>${esc(m.dest)}</td><td class="mono">${esc(m.sos.join(', '))}</td><td class="mono">${esc(m.art)}</td><td class="num">${m.qty}</td><td>${regs.has(code(m.dest)) ? 'Sí' : '<span class="det">Fuera de región</span>'}</td><td>${esc(m.added)}</td><td><button class="btn ghost sm" data-del="${i}">Quitar</button></td></tr>`).join('') + '</tbody>'
    : '<tbody><tr><td class="det">No hay TOs manuales.</td></tr></tbody>';
}

function renderOut() {
  $('#cntOut').textContent = S.out.length || '';
  $('#tblOut').innerHTML = S.out.length ? `<thead><tr><th>SO</th><th>Agencia</th><th>Artículo</th><th>Estatus pago</th><th>Último estatus</th><th>Comentario</th><th>Detectado</th></tr></thead><tbody>` +
    S.out.map(x => `<tr><td class="mono">${esc(x.so)}</td><td>${esc(x.ub)}</td><td class="mono">${esc(x.art)}</td><td>${esc(x.pago)}</td><td>${x.cls ? `<span class="badge ${CLS[x.cls].b}">${CLS[x.cls].t}</span>` : ''}</td><td>${x.com?.v ? esc(COM_BY[x.com.v]?.t || x.com.v) + (x.com.texto ? ' · ' + esc(x.com.texto) : '') : ''}</td><td>${esc(x.salio)}</td></tr>`).join('') + '</tbody>'
    : '<tbody><tr><td class="det">Nada todavía. Aparecerán aquí las SOs que estaban en la carga anterior y ya no vienen en el archivo nuevo.</td></tr></tbody>';
}

// ---------- Exportar / respaldo ----------
function exportXlsx() {
  if (!S.result.length) return toast('No hay datos para exportar');
  const rows = filtered().map(r => {
    const c = S.comments[r.so] || {};
    return {
      Prioridad: r.prio, Agencia: r.ub, Departamento: r.dep, Cedis: r.cedis === '15002' ? 'TUL' : 'CHA', SO: r.so, Artículo: r.art, Cantidad: r.qty,
      'Estatus pago': r.pago, Fecha: r.fecha, Antigüedad: r.ant, 'TO asignada': r.toA, Clasificación: CLS[r.cls].t, Detalle: r.det,
      Alertas: r.alerts.join(' | '), 'Días en estatus': diasEst(r), Comentario: COM_BY[c.v]?.t || '', 'Fecha compromiso': c.fecha || '', 'Nota': c.texto || '',
    };
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0] || {}).map(k => ({ wch: Math.min(40, Math.max(k.length + 2, 10)) }));
  XLSX.utils.book_append_sheet(wb, ws, 'Seguimiento SO');
  XLSX.writeFile(wb, `Monitor_SO_${today()}.xlsx`);
}
function backup() {
  const blob = new Blob([JSON.stringify({ v: 1, at: new Date().toISOString(), comments: S.comments, hist: S.hist, manual: S.manual, out: S.out, lastSOs: S.lastSOs, lastSig: S.lastSig, data: S.data }, null, 1)], { type: 'application/json' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `Monitor_SO_respaldo_${today()}.json`; a.click(); URL.revokeObjectURL(a.href);
}
async function restore(file) {
  try {
    const j = JSON.parse(await file.text());
    if (j.comments) S.comments = { ...S.comments, ...j.comments };
    if (j.hist) S.hist = { ...j.hist, ...S.hist };
    if (j.manual) { const seen = new Set(S.manual.map(m => m.to + '|' + m.art)); j.manual.forEach(m => { if (!seen.has(m.to + '|' + m.art)) S.manual.push(m); }); }
    if (j.data && !S.data.so) S.data = j.data;
    if (j.out && !S.out.length) S.out = j.out;
    if (j.lastSOs && !S.lastSOs.length) { S.lastSOs = j.lastSOs; S.lastSig = j.lastSig || ''; }
    ['comments', 'hist', 'manual', 'data', 'out', 'lastSOs', 'lastSig'].forEach(k => LS.set(k, S[k]));
    purgeManual(); run(false); toast('Respaldo importado');
  } catch (e) { toast('Respaldo inválido: ' + e.message); }
}

// ---------- Eventos ----------
function bind() {
  const inF = $('#inFiles'), drop = $('#drop');
  $('#btnPick').onclick = () => inF.click();
  inF.onchange = () => { readFiles([...inF.files]); inF.value = ''; };
  drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); readFiles([...e.dataTransfer.files]); };

  $('#tabs').onclick = e => {
    const b = e.target.closest('button'); if (!b) return;
    $$('#tabs button').forEach(x => x.classList.toggle('on', x === b));
    $$('.tab').forEach(t => t.classList.toggle('hidden', t.id !== b.dataset.tab));
  };
  const showSO = () => $('#tabs button[data-tab=tabSO]').click();

  $('#kpis').onclick = e => {
    const k = e.target.closest('.kpi'); if (!k) return;
    S.f.cls = new Set(k.dataset.k ? [k.dataset.k] : []); renderKpis(); renderFilters(); renderTable(); showSO();
  };
  $('#fCls').onclick = e => { const c = e.target.closest('[data-cls]'); if (!c) return; const v = c.dataset.cls; S.f.cls.has(v) ? S.f.cls.delete(v) : S.f.cls.add(v); renderKpis(); renderFilters(); renderTable(); };
  $('#fPay').onclick = e => { const c = e.target.closest('[data-pay]'); if (!c) return; const v = c.dataset.pay; S.f.pay.has(v) ? S.f.pay.delete(v) : S.f.pay.add(v); renderFilters(); renderTable(); };
  $('#fAg').onchange = e => { S.f.ag = e.target.value; renderTable(); };
  $('#fCom').onchange = e => { S.f.com = e.target.value; renderTable(); };
  $('#fAlert').onchange = e => { S.f.alert = e.target.checked; renderTable(); };
  $('#fSearch').oninput = e => { S.f.q = e.target.value; renderTable(); };
  $('#fClear').onclick = () => { S.f = { cls: new Set(), pay: new Set(), ag: '', q: '', com: '', alert: false }; $('#fSearch').value = ''; $('#fAlert').checked = false; renderKpis(); renderFilters(); renderTable(); };

  $('#tblSO thead').onclick = e => { const th = e.target.closest('th'); if (!th) return; const k = th.dataset.sort; S.sort = { k, dir: S.sort.k === k ? -S.sort.dir : 1 }; renderTable(); };
  $('#tblSO tbody').addEventListener('change', e => {
    const f = e.target.dataset.com; if (!f) return;
    const so = e.target.closest('tr').dataset.so;
    const c = S.comments[so] || {};
    c[f] = e.target.value;
    if (f === 'v' && !c.v) delete S.comments[so]; else { c.at = today(); S.comments[so] = c; }
    LS.set('comments', S.comments);
    if (f === 'v') renderTable();
    toast('Comentario guardado');
  });

  $('#tblAg').onclick = e => {
    const a = e.target.closest('a[data-ag]'); if (!a) return; e.preventDefault();
    S.f.ag = a.dataset.ag; S.f.cls = new Set(a.dataset.cls ? [a.dataset.cls] : []);
    renderKpis(); renderFilters(); renderTable(); showSO();
  };

  $('#btnManAdd').onclick = () => {
    const add = parseManual($('#manPaste').value);
    if (!add.length) { $('#manMsg').textContent = 'No encontré líneas válidas (deben empezar con TO…).'; return; }
    const seen = new Set(S.manual.map(m => m.to + '|' + m.art));
    let n = 0; add.forEach(m => { if (!seen.has(m.to + '|' + m.art)) { S.manual.push(m); seen.add(m.to + '|' + m.art); n++; } });
    LS.set('manual', S.manual); purgeManual();
    $('#manMsg').textContent = `${n} línea(s) agregadas${add.length - n ? `, ${add.length - n} repetidas` : ''}.`;
    $('#manPaste').value = ''; run(false);
  };
  $('#btnManClear').onclick = () => { if (confirm('¿Borrar todas las TOs manuales?')) { S.manual = []; LS.set('manual', S.manual); run(false); } };
  $('#tblMan').onclick = e => { const b = e.target.closest('[data-del]'); if (!b) return; S.manual.splice(+b.dataset.del, 1); LS.set('manual', S.manual); run(false); };

  $('#btnExport').onclick = exportXlsx;
  $('#btnBackup').onclick = backup;
  $('#inBackup').onchange = e => { if (e.target.files[0]) restore(e.target.files[0]); e.target.value = ''; };
}

bind();
run(false);
