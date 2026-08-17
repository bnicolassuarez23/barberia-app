import React, { useState, useEffect, useMemo } from 'react';
import {
  Plus, Trash2, ChevronLeft, ChevronRight, Scissors, Shirt, Gift,
  CalendarDays, ListChecks, Users, Loader2, Wallet, TrendingUp, TrendingDown,
  Package, ShoppingCart, AlertTriangle, ArrowRightLeft
} from 'lucide-react';
import { supabase } from './supabaseClient';
import Auth from './Auth';
import ExcelJS from 'exceljs';

const TIPOS = [
  { key: 'corte', label: 'Corte', Icon: Scissors },
  { key: 'ropa', label: 'Ropa/Producto', Icon: Shirt },
  { key: 'propina', label: 'Propina', Icon: Gift },
];
const METODOS = [
  { key: 'efectivo', label: 'Efectivo' },
  { key: 'qr', label: 'QR/Transf.' },
  { key: 'tarjeta', label: 'Tarjeta' },
];
const GASTO_CATEGORIAS = ['Alquiler', 'Insumos/Productos', 'Servicios', 'Sueldos', 'Otro'];
const DIAS_LARGOS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS_SEMANA_CORTA = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];

function pad2(n) { return String(n).padStart(2, '0'); }
function toISO(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }
function fromISO(s) { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); }
function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function mondayOf(d) { const day = d.getDay(); const diff = day === 0 ? -6 : 1 - day; return addDays(d, diff); }
function formatLong(d) { return `${DIAS_LARGOS[d.getDay()]} ${d.getDate()} de ${MESES[d.getMonth()]} de ${d.getFullYear()}`; }
function formatShort(d) { return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}`; }
function money(n) { return Math.round(n || 0).toLocaleString('es-AR'); }
function nowHHMM() { const d = new Date(); return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; }
function uid() { return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`; }

const FONT_DISPLAY = "'Fraunces', serif";
const FONT_BODY = "'IBM Plex Sans', sans-serif";
const FONT_MONO = "'IBM Plex Mono', monospace";

function Sello({ children, tone = 'alto' }) {
  const cls = tone === 'alto' ? 'border-amber-500 text-amber-400' : 'border-stone-600 text-stone-500';
  return (
    <span className={`inline-block border-2 ${cls} px-2 py-0.5 rounded -rotate-2 text-xs uppercase tracking-widest font-semibold`} style={{ fontFamily: FONT_MONO }}>
      {children}
    </span>
  );
}

function BarraBarbero() {
  return (
    <div
      className="h-1.5 w-full"
      style={{ backgroundImage: 'repeating-linear-gradient(-45deg, #b91c1c 0px, #b91c1c 10px, #f5f5f4 10px, #f5f5f4 20px, #1d4ed8 20px, #1d4ed8 30px)' }}
    />
  );
}

function Fila({ label, valor, sub, destacado }) {
  return (
    <div className={`flex items-baseline justify-between ${sub ? 'pl-3 text-stone-400 text-sm' : ''}`}>
      <span className={destacado ? 'text-stone-200' : ''}>{label}</span>
      <span
        className={destacado === 'amber' ? 'text-amber-400 font-semibold' : destacado === 'stone' ? 'text-stone-200 font-semibold' : ''}
        style={{ fontFamily: FONT_MONO }}
      >
        ${money(valor)}
      </span>
    </div>
  );
}

function StatCard({ label, valor, tono, esNumero }) {
  return (
    <div className="bg-stone-900 border border-stone-800 rounded-lg p-3">
      <p className="text-xs text-stone-500 mb-1">{label}</p>
      <p className={`text-lg font-semibold ${tono === 'amber' ? 'text-amber-400' : tono === 'red' ? 'text-red-400' : tono === 'emerald' ? 'text-emerald-400' : 'text-stone-100'}`} style={{ fontFamily: FONT_MONO }}>
        {esNumero ? valor : `$${money(valor)}`}
      </p>
    </div>
  );
}

function BarberoDia({ barbero, pct, registros = [], onAdd, onDelete, fecha }) {
  const [monto, setMonto] = useState('');
  const [tipo, setTipo] = useState('corte');
  const [metodo, setMetodo] = useState('efectivo');
  const [hora, setHora] = useState(nowHHMM());

  const subtotalCortes = registros.filter((r) => r.tipo === 'corte').reduce((s, r) => s + r.monto, 0);
  const propinas = registros.filter((r) => r.tipo === 'propina').reduce((s, r) => s + r.monto, 0);

  const bNombre = barbero || '';
  const esJefe = bNombre.toLowerCase() === 'nico' || bNombre.toLowerCase() === 'martín' || bNombre.toLowerCase() === 'martin';
  const pctEfectivo = esJefe ? 100 : pct;

  const comision = subtotalCortes * (pctEfectivo / 100);
  const aCobrar = comision + propinas;

  const submit = () => {
    const m = parseFloat(monto);
    if (!m || m <= 0) return;

    if (onAdd) {
      onAdd({
        barbero,
        tipo,
        metodo,
        hora,
        monto: m,
        fecha: fecha
      });
    }

    setMonto('');
    setHora(nowHHMM());
  };

  return (
    <div className="bg-stone-900 border border-stone-800 rounded-lg overflow-hidden">
      <div className="flex items-center justify-between px-4 py-3 border-b border-stone-800">
        <h3 className="font-semibold text-lg" style={{ fontFamily: FONT_DISPLAY }}>{barbero}</h3>
        {registros.length > 0 && <span className="text-amber-400 text-sm" style={{ fontFamily: FONT_MONO }}>${money(aCobrar)}</span>}
      </div>

      <div className="px-4 py-3 space-y-2">
        <div className="grid grid-cols-2 gap-2">
          <input
            type="number" inputMode="numeric" placeholder="Monto $" value={monto}
            onChange={(e) => setMonto(e.target.value)}
            className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600"
            style={{ fontFamily: FONT_MONO }}
          />
          <input
            type="time" value={hora} onChange={(e) => setHora(e.target.value)}
            className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100"
            style={{ fontFamily: FONT_MONO }}
          />
        </div>
        <div className="grid grid-cols-2 gap-2">
          <select value={tipo} onChange={(e) => setTipo(e.target.value)} className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100">
            {TIPOS.map((t) => <option key={t.key} value={t.key}>{t.label}</option>)}
          </select>
          <select value={metodo} onChange={(e) => setMetodo(e.target.value)} className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100">
            {METODOS.map((m) => <option key={m.key} value={m.key}>{m.label}</option>)}
          </select>
        </div>
        <button onClick={submit} className="w-full flex items-center justify-center gap-1 bg-amber-500 hover:bg-amber-400 active:bg-amber-600 text-stone-950 font-semibold rounded py-2">
          <Plus size={16} /> Agregar
        </button>
      </div>

      {registros.length > 0 && (
        <div className="border-t border-stone-800 divide-y divide-stone-800">
          {registros.slice().sort((a, b) => (a.hora || '').localeCompare(b.hora || '')).map((r) => {
            const T = TIPOS.find((t) => t.key === r.tipo);
            return (
              <div key={r.id} className="flex items-center justify-between px-4 py-2 text-sm">
                <div className="flex items-center gap-2 text-stone-400">
                  {T && <T.Icon size={14} />}
                  <span style={{ fontFamily: FONT_MONO }}>{r.hora}</span>
                  <span className="text-stone-600">·</span>
                  <span>{METODOS.find((m) => m.key === r.metodo)?.label}</span>
                </div>
                <div className="flex items-center gap-3">
                  <span style={{ fontFamily: FONT_MONO }}>${money(r.monto)}</span>
                  <button onClick={() => onDelete(r.id)} className="text-stone-600 active:text-red-400"><Trash2 size={14} /></button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function HoyView({ dateSel, setDateSel, barberos = [], pct, registros = [], addRegistro, deleteRegistro }) {
  const fecha = fromISO(dateSel);
  const delDia = useMemo(() => (registros || []).filter((r) => r.fecha === dateSel), [registros, dateSel]);
  const cambiarDia = (n) => setDateSel(toISO(addDays(fecha, n)));

  const totales = useMemo(() => {
    let facturado = 0, efectivo = 0, qr = 0, tarjeta = 0, aPagarBarberos = 0, paraBarberia = 0;
    delDia.forEach((r) => {
      const barberoStr = r.barbero || '';
      const esJefe = barberoStr.toLowerCase() === 'nico' || barberoStr.toLowerCase() === 'martín' || barberoStr.toLowerCase() === 'martin';
      const pctEfectivo = esJefe ? 100 : pct;

      if (r.tipo === 'corte') {
        facturado += r.monto;
        aPagarBarberos += r.monto * (pctEfectivo / 100);
        paraBarberia += r.monto * (1 - pctEfectivo / 100);
      } else if (r.tipo === 'ropa') {
        facturado += r.monto;
        paraBarberia += r.monto;
      } else if (r.tipo === 'propina') {
        aPagarBarberos += r.monto;
      }
      if (r.metodo === 'efectivo') efectivo += r.monto;
      else if (r.metodo === 'qr') qr += r.monto;
      else if (r.metodo === 'tarjeta') tarjeta += r.monto;
    });
    return { facturado, efectivo, qr, tarjeta, aPagarBarberos, paraBarberia };
  }, [delDia, pct]);

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <button onClick={() => cambiarDia(-1)} className="p-2 text-stone-400 active:text-amber-400"><ChevronLeft size={20} /></button>
        <div className="text-center">
          <p className="capitalize text-sm text-stone-300">{formatLong(fecha)}</p>
          {dateSel !== toISO(new Date()) && (
            <button onClick={() => setDateSel(toISO(new Date()))} className="text-xs text-amber-400 mt-0.5" style={{ fontFamily: FONT_MONO }}>volver a hoy</button>
          )}
        </div>
        <button onClick={() => cambiarDia(1)} className="p-2 text-stone-400 active:text-amber-400"><ChevronRight size={20} /></button>
      </div>

      {barberos.length === 0 && (
        <p className="text-stone-500 text-sm">Todavía no agregaste barberos. Andá a la pestaña "Barberos" para sumar el primero.</p>
      )}

      {barberos.map((b) => {
        const nombreBarbero = typeof b === 'string' ? b : (b?.nombre || '');
        return (
          <BarberoDia
            key={nombreBarbero}
            barbero={nombreBarbero}
            pct={typeof b === 'object' ? b.pct : pct}
            registros={delDia.filter((r) => (r?.barbero || '') === nombreBarbero)}
            fecha={dateSel}
            onAdd={addRegistro}
            onDelete={deleteRegistro}
          />
        );
      })}

      {delDia.length > 0 && (
        <div className="bg-stone-900 border border-stone-800 rounded-lg p-4 space-y-2">
          <p className="text-xs uppercase tracking-widest text-stone-500" style={{ fontFamily: FONT_MONO }}>Resumen del día</p>
          <Fila label="Total facturado" valor={totales.facturado} />
          <Fila label="Efectivo" valor={totales.efectivo} sub />
          <Fila label="QR/Transferencia" valor={totales.qr} sub />
          <Fila label="Tarjeta" valor={totales.tarjeta} sub />
          <div className="border-t border-dashed border-stone-700 my-1" />
          <Fila label="A pagar a barberos" valor={totales.aPagarBarberos} destacado="amber" />
          <Fila label="Para la barbería" valor={totales.paraBarberia} destacado="stone" />
        </div>
      )}
    </div>
  );
}

function CalendarioView({ monthAnchor, setMonthAnchor, registros = [], onSelectDay, dateSel }) {
  const year = monthAnchor.getFullYear();
  const month = monthAnchor.getMonth();
  const first = new Date(year, month, 1);
  const startOffset = (first.getDay() + 6) % 7;
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  const totalesPorDia = useMemo(() => {
    const map = {};
    (registros || []).forEach((r) => {
      if (r.tipo === 'corte' || r.tipo === 'ropa') map[r.fecha] = (map[r.fecha] || 0) + r.monto;
    });
    return map;
  }, [registros]);

  const celdas = [];
  for (let i = 0; i < startOffset; i++) celdas.push(null);
  for (let d = 1; d <= daysInMonth; d++) celdas.push(d);

  const cambiarMes = (n) => setMonthAnchor(new Date(year, month + n, 1));

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <button onClick={() => cambiarMes(-1)} className="p-2 text-stone-400 active:text-amber-400"><ChevronLeft size={20} /></button>
        <h2 className="capitalize text-lg" style={{ fontFamily: FONT_DISPLAY }}>{MESES[month]} {year}</h2>
        <button onClick={() => cambiarMes(1)} className="p-2 text-stone-400 active:text-amber-400"><ChevronRight size={20} /></button>
      </div>
      <div className="grid grid-cols-7 gap-1 text-center text-xs text-stone-500" style={{ fontFamily: FONT_MONO }}>
        {DIAS_SEMANA_CORTA.concat('Dom').map((d) => <div key={d}>{d}</div>)}
      </div>
      <div className="grid grid-cols-7 gap-1">
        {celdas.map((d, i) => {
          if (d === null) return <div key={`b${i}`} />;
          const iso = toISO(new Date(year, month, d));
          const esDomingo = new Date(year, month, d).getDay() === 0;
          const total = totalesPorDia[iso] || 0;
          const esHoy = iso === toISO(new Date());
          const esSel = iso === dateSel;
          return (
            <button
              key={iso} onClick={() => onSelectDay(iso)}
              className={`h-12 rounded flex flex-col items-center justify-center border ${esSel ? 'border-amber-500' : 'border-stone-800'} ${esDomingo ? 'bg-stone-900/40' : 'bg-stone-900'}`}
            >
              <span className={`text-sm ${esHoy ? 'text-amber-400 font-semibold' : esDomingo ? 'text-stone-600' : 'text-stone-200'}`}>{d}</span>
              {total > 0 && (
                <span className="text-xs text-stone-500" style={{ fontFamily: FONT_MONO }}>{total >= 1000 ? `${Math.round(total / 1000)}k` : total}</span>
              )}
            </button>
          );
        })}
      </div>
      <p className="text-xs text-stone-600 text-center">Tocá un día para ver o cargar sus registros.</p>
    </div>
  );
}

function CierresView({ weekAnchor, setWeekAnchor, registros = [], barberos = [], pct, gastos = [] }) {
  const lunes = weekAnchor;
  const sabado = addDays(lunes, 5);
  const isoLunes = toISO(lunes);
  const isoSabado = toISO(sabado);

  const regsSemana = useMemo(() => (registros || []).filter((r) => r.fecha >= isoLunes && r.fecha <= isoSabado), [registros, isoLunes, isoSabado]);
  const gastosSemana = useMemo(() => (gastos || []).filter((g) => g.fecha >= isoLunes && g.fecha <= isoSabado), [gastos, isoLunes, isoSabado]);
  const totalGastosSemana = gastosSemana.reduce((s, g) => s + g.monto, 0);
  const cambiarSemana = (n) => setWeekAnchor(addDays(lunes, n * 7));

  const stats = useMemo(() => {
    const cortes = regsSemana.filter((r) => r.tipo === 'corte');
    const ropas = regsSemana.filter((r) => r.tipo === 'ropa');
    const propinas = regsSemana.filter((r) => r.tipo === 'propina');

    const totalCortes = cortes.reduce((s, r) => s + r.monto, 0);
    const totalRopa = ropas.reduce((s, r) => s + r.monto, 0);
    const totalPropinas = propinas.reduce((s, r) => s + r.monto, 0);
    const facturado = totalCortes + totalRopa;
    const aPagarBarberos = totalCortes * (pct / 100) + totalPropinas;
    const paraBarberia = totalCortes * (1 - pct / 100) + totalRopa;

    const porBarbero = (barberos || []).map((b) => {
      const nombreB = typeof b === 'string' ? b : (b?.nombre || '');
      const cb = cortes.filter((r) => r.barbero === nombreB);
      const totalB = cb.reduce((s, r) => s + r.monto, 0);
      const propB = propinas.filter((r) => r.barbero === nombreB).reduce((s, r) => s + r.monto, 0);
      const ropaB = ropas.filter((r) => r.barbero === nombreB).reduce((s, r) => s + r.monto, 0);
      const esJefe = nombreB.toLowerCase() === 'nico' || nombreB.toLowerCase() === 'martín' || nombreB.toLowerCase() === 'martin';
      const pctEfectivo = esJefe ? 100 : pct;
      const comisionB = totalB * (pctEfectivo / 100);
      return { barbero: nombreB, cantidad: cb.length, totalB, comisionB, propB, ropaB, aCobrar: comisionB + propB };
    });

    const porDia = [0, 1, 2, 3, 4, 5].map((i) => {
      const iso = toISO(addDays(lunes, i));
      const regsDia = cortes.filter((r) => r.fecha === iso);
      return { iso, cantidad: regsDia.length, monto: regsDia.reduce((s, r) => s + r.monto, 0) };
    });

    const horasMap = {};
    cortes.forEach((r) => {
      const h = (r.hora || '00:00').split(':')[0];
      horasMap[h] = (horasMap[h] || 0) + 1;
    });
    const porHora = Object.entries(horasMap).map(([hora, cantidad]) => ({ hora, cantidad })).sort((a, b) => a.hora.localeCompare(b.hora));

    const efectivo = regsSemana.filter((r) => r.metodo === 'efectivo').reduce((s, r) => s + r.monto, 0);
    const qr = regsSemana.filter((r) => r.metodo === 'qr').reduce((s, r) => s + r.monto, 0);
    const tarjeta = regsSemana.filter((r) => r.metodo === 'tarjeta').reduce((s, r) => s + r.monto, 0);

    return { totalCortes, totalRopa, totalPropinas, facturado, aPagarBarberos, paraBarberia, porBarbero, porDia, porHora, efectivo, qr, tarjeta, cantidadCortes: cortes.length };
  }, [regsSemana, barberos, pct, lunes]);

  const diaMax = stats.porDia.reduce((max, d) => (d.cantidad > max.cantidad ? d : max), stats.porDia[0]);
  const diaMin = stats.porDia.reduce((min, d) => (d.cantidad < min.cantidad ? d : min), stats.porDia[0]);
  const hayVariacionDia = diaMax && diaMin && diaMax.cantidad > diaMin.cantidad;

  const horaMax = stats.porHora.length ? stats.porHora.reduce((max, h) => (h.cantidad > max.cantidad ? h : max), stats.porHora[0]) : null;
  const horaMin = stats.porHora.length ? stats.porHora.reduce((min, h) => (h.cantidad < min.cantidad ? h : min), stats.porHora[0]) : null;
  const hayVariacionHora = horaMax && horaMin && horaMax.cantidad > horaMin.cantidad;

  const maxDiaCantidad = Math.max(1, ...stats.porDia.map((d) => d.cantidad));
  const maxHoraCantidad = Math.max(1, ...stats.porHora.map((h) => h.cantidad));
  const gananciaNeta = stats.paraBarberia - totalGastosSemana;

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <button onClick={() => cambiarSemana(-1)} className="p-2 text-stone-400 active:text-amber-400"><ChevronLeft size={20} /></button>
        <div className="text-center">
          <p className="text-xs uppercase tracking-widest text-stone-500" style={{ fontFamily: FONT_MONO }}>Semana</p>
          <p style={{ fontFamily: FONT_DISPLAY }}>{formatShort(lunes)} — {formatShort(sabado)}</p>
        </div>
        <button onClick={() => cambiarSemana(1)} className="p-2 text-stone-400 active:text-amber-400"><ChevronRight size={20} /></button>
      </div>

      <div className="grid grid-cols-2 gap-3">
        <StatCard label="Total facturado" valor={stats.facturado} />
        <StatCard label="Cortes realizados" valor={stats.cantidadCortes} esNumero />
        <StatCard label="A pagar a barberos" valor={stats.aPagarBarberos} tono="amber" />
        <StatCard label="Para la barbería" valor={stats.paraBarberia} />
      </div>

      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Ganancia neta</h3>
        <div className="bg-stone-900 border border-stone-800 rounded-lg p-4 space-y-2">
          <Fila label="Para la barbería (cortes + ropa)" valor={stats.paraBarberia} />
          <Fila label="Gastos del local" valor={totalGastosSemana} />
          <div className="border-t border-dashed border-stone-700 my-1" />
          <div className="flex items-baseline justify-between">
            <span className="text-stone-200">Ganancia neta</span>
            <span className={`font-semibold ${gananciaNeta >= 0 ? 'text-emerald-400' : 'text-red-400'}`} style={{ fontFamily: FONT_MONO }}>
              ${money(gananciaNeta)}
            </span>
          </div>
        </div>
      </section>

      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Resumen por barbero</h3>
        <div className="space-y-2">
          {stats.porBarbero.map((b) => (
            <div key={b.barbero} className="bg-stone-900 border border-stone-800 rounded-lg p-3">
              <div className="flex items-center justify-between mb-1">
                <span className="font-semibold" style={{ fontFamily: FONT_DISPLAY }}>{b.barbero}</span>
                <span className="text-amber-400" style={{ fontFamily: FONT_MONO }}>${money(b.aCobrar)}</span>
              </div>
              <div className="text-xs text-stone-400 space-y-0.5" style={{ fontFamily: FONT_MONO }}>
                <div className="flex justify-between"><span>{b.cantidad} cortes · ${money(b.totalB)}</span><span>comisión ${money(b.comisionB)}</span></div>
                <div className="flex justify-between"><span>propinas ${money(b.propB)}</span><span>ropa vendida ${money(b.ropaB)}</span></div>
              </div>
            </div>
          ))}
          {barberos.length === 0 && <p className="text-stone-500 text-sm">No hay barberos cargados.</p>}
        </div>
      </section>

      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Actividad por día</h3>
        <div className="bg-stone-900 border border-stone-800 rounded-lg p-4">
          <div className="flex items-end gap-2 h-28">
            {stats.porDia.map((d, i) => (
              <div key={d.iso} className="flex-1 flex flex-col items-center justify-end gap-1 h-full">
                <span className="text-xs text-stone-400" style={{ fontFamily: FONT_MONO }}>{d.cantidad}</span>
                <div
                  className={`w-full rounded-t ${hayVariacionDia && d.iso === diaMax.iso ? 'bg-amber-500' : hayVariacionDia && d.iso === diaMin.iso ? 'bg-stone-700' : 'bg-stone-600'}`}
                  style={{ height: `${Math.max(6, (d.cantidad / maxDiaCantidad) * 100)}%` }}
                />
                <span className="text-xs text-stone-500">{DIAS_SEMANA_CORTA[i]}</span>
              </div>
            ))}
          </div>
          {hayVariacionDia && (
            <div className="flex justify-between mt-3">
              <Sello tone="alto">Pico: {DIAS_SEMANA_CORTA[stats.porDia.indexOf(diaMax)]}</Sello>
              <Sello tone="bajo">Bajo: {DIAS_SEMANA_CORTA[stats.porDia.indexOf(diaMin)]}</Sello>
            </div>
          )}
        </div>
      </section>

      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Actividad por hora</h3>
        {stats.porHora.length === 0 ? (
          <p className="text-stone-500 text-sm">Todavía no hay cortes registrados con horario esta semana.</p>
        ) : (
          <div className="bg-stone-900 border border-stone-800 rounded-lg p-4 space-y-2">
            {stats.porHora.map((h) => (
              <div key={h.hora} className="flex items-center gap-3">
                <span className="text-xs text-stone-400 w-10" style={{ fontFamily: FONT_MONO }}>{h.hora}h</span>
                <div className="flex-1 bg-stone-800 rounded h-3 overflow-hidden">
                  <div
                    className={`h-full ${hayVariacionHora && h.hora === horaMax.hora ? 'bg-amber-500' : hayVariacionHora && h.hora === horaMin.hora ? 'bg-stone-600' : 'bg-stone-500'}`}
                    style={{ width: `${Math.max(6, (h.cantidad / maxHoraCantidad) * 100)}%` }}
                  />
                </div>
                <span className="text-xs text-stone-400 w-6 text-right" style={{ fontFamily: FONT_MONO }}>{h.cantidad}</span>
              </div>
            ))}
            {hayVariacionHora && (
              <div className="flex justify-between pt-2">
                <Sello tone="alto">Pico: {horaMax.hora}h</Sello>
                <Sello tone="bajo">Bajo: {horaMin.hora}h</Sello>
              </div>
            )}
          </div>
        )}
      </section>

      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Métodos de pago</h3>
        <div className="bg-stone-900 border border-stone-800 rounded-lg p-4 space-y-2">
          <Fila label="Efectivo" valor={stats.efectivo} />
          <Fila label="QR/Transferencia" valor={stats.qr} />
          <Fila label="Tarjeta" valor={stats.tarjeta} />
        </div>
      </section>
    </div>
  );
}

function BarberosView({ config, onSave }) {
  const [nuevo, setNuevo] = useState('');
  const [pctInput, setPctInput] = useState(String(config?.comisionPct || 50));

  const agregar = () => {
    const nombre = nuevo.trim();
    if (!nombre || config.barberos.includes(nombre)) return;
    onSave({ ...config, barberos: [...config.barberos, nombre] });
    setNuevo('');
  };

  const quitar = (nombre) => onSave({ ...config, barberos: config.barberos.filter((b) => b !== nombre) });

  const guardarPct = () => {
    const p = parseFloat(pctInput);
    if (isNaN(p) || p < 0 || p > 100) return;
    onSave({ ...config, comisionPct: p });
  };

  return (
    <div className="space-y-6">
      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Barberos activos</h3>
        <div className="space-y-2">
          {(config?.barberos || []).map((b) => (
            <div key={b} className="flex items-center justify-between bg-stone-900 border border-stone-800 rounded-lg px-4 py-2.5">
              <span style={{ fontFamily: FONT_DISPLAY }}>{b}</span>
              <button onClick={() => quitar(b)} className="text-stone-600 active:text-red-400"><Trash2 size={16} /></button>
            </div>
          ))}
          {(config?.barberos || []).length === 0 && <p className="text-stone-500 text-sm">No hay barberos cargados todavía.</p>}
        </div>
        <div className="flex gap-2 mt-3">
          <input
            value={nuevo} onChange={(e) => setNuevo(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && agregar()}
            placeholder="Nombre del barbero"
            className="flex-1 bg-stone-900 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600"
          />
          <button onClick={agregar} className="bg-amber-500 active:bg-amber-600 text-stone-950 font-semibold rounded px-4"><Plus size={18} /></button>
        </div>
        <p className="text-xs text-stone-600 mt-2">Sacar a un barbero no borra sus registros anteriores, solo deja de aparecer para cargar cortes nuevos.</p>
      </section>

      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Comisión por corte</h3>
        <div className="bg-stone-900 border border-stone-800 rounded-lg p-4">
          <p className="text-sm text-stone-400 mb-3">Porcentaje que se queda cada barbero por cada corte. Las propinas son 100% del barbero y la ropa/productos 100% de la barbería.</p>
          <div className="flex gap-2 items-center">
            <input
              type="number" value={pctInput} onChange={(e) => setPctInput(e.target.value)}
              className="w-24 bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100" style={{ fontFamily: FONT_MONO }}
            />
            <span className="text-stone-400">%</span>
            <button onClick={guardarPct} className="ml-auto bg-amber-500 active:bg-amber-600 text-stone-950 font-semibold rounded px-4 py-2">Guardar</button>
          </div>
        </div>
      </section>
    </div>
  );
}

function GastosView({ gastos = [], onAdd, onDelete }) {
  const [monto, setMonto] = useState('');
  const [categoria, setCategoria] = useState(GASTO_CATEGORIAS[0]);
  const [fecha, setFecha] = useState(toISO(new Date()));
  const [descripcion, setDescripcion] = useState('');
  const [mesAnchor, setMesAnchor] = useState(new Date());

  const submit = () => {
    const m = parseFloat(monto);
    if (!m || m <= 0) return;

    if (onAdd) {
      onAdd({
        monto: m,
        categoria,
        fecha,
        descripcion: descripcion.trim()
      });
    }

    setMonto('');
    setDescripcion('');
  };

  const totalGeneral = (gastos || []).reduce((s, g) => s + g.monto, 0);

  const mesAnchorISO = `${mesAnchor.getFullYear()}-${pad2(mesAnchor.getMonth() + 1)}`;
  const cambiarMes = (n) => setMesAnchor(new Date(mesAnchor.getFullYear(), mesAnchor.getMonth() + n, 1));
  const esMesActual = mesAnchorISO === toISO(new Date()).slice(0, 7);

  const gastosDelMes = useMemo(() => {
    return (gastos || [])
      .filter((g) => (g.fecha || '').slice(0, 7) === mesAnchorISO)
      .sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
  }, [gastos, mesAnchorISO]);

  const totalMes = gastosDelMes.reduce((s, g) => s + g.monto, 0);

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-2 gap-3">
        <StatCard label="Gastos del mes" valor={totalMes} />
        <StatCard label="Gastos totales" valor={totalGeneral} />
      </div>

      <div className="bg-stone-900 border border-stone-800 rounded-lg p-4 space-y-2">
        <p className="text-xs uppercase tracking-widest text-stone-500" style={{ fontFamily: FONT_MONO }}>Nuevo gasto</p>
        <div className="grid grid-cols-2 gap-2">
          <input
            type="number" inputMode="numeric" placeholder="Monto $" value={monto}
            onChange={(e) => setMonto(e.target.value)}
            className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600"
            style={{ fontFamily: FONT_MONO }}
          />
          <input
            type="date" value={fecha} onChange={(e) => setFecha(e.target.value)}
            className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100"
            style={{ fontFamily: FONT_MONO }}
          />
        </div>
        <select value={categoria} onChange={(e) => setCategoria(e.target.value)} className="w-full bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100">
          {GASTO_CATEGORIAS.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
        <input
          value={descripcion} onChange={(e) => setDescripcion(e.target.value)}
          placeholder="Descripción (opcional)"
          className="w-full bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600"
        />
        <button onClick={submit} className="w-full flex items-center justify-center gap-1 bg-amber-500 hover:bg-amber-400 active:bg-amber-600 text-stone-950 font-semibold rounded py-2">
          <Plus size={16} /> Agregar gasto
        </button>
      </div>

      <div className="flex items-center justify-between">
        <button onClick={() => cambiarMes(-1)} className="p-2 text-stone-400 active:text-amber-400"><ChevronLeft size={20} /></button>
        <div className="text-center">
          <p className="capitalize text-sm text-stone-300" style={{ fontFamily: FONT_MONO }}>{MESES[mesAnchor.getMonth()]} {mesAnchor.getFullYear()}</p>
          {!esMesActual && (
            <button onClick={() => setMesAnchor(new Date())} className="text-xs text-amber-400 mt-0.5" style={{ fontFamily: FONT_MONO }}>volver a este mes</button>
          )}
        </div>
        <button onClick={() => cambiarMes(1)} className="p-2 text-stone-400 active:text-amber-400"><ChevronRight size={20} /></button>
      </div>

      {gastosDelMes.length === 0 && <p className="text-stone-500 text-sm">No hay gastos cargados en {MESES[mesAnchor.getMonth()]}.</p>}

      {gastosDelMes.length > 0 && (
        <div className="bg-stone-900 border border-stone-800 rounded-lg divide-y divide-stone-800">
          {gastosDelMes.map((g) => (
            <div key={g.id} className="flex items-center justify-between px-4 py-2.5 text-sm">
              <div>
                <div className="flex items-center gap-2">
                  <span style={{ fontFamily: FONT_MONO }}>{formatShort(fromISO(g.fecha))}</span>
                  <span className="text-stone-500">·</span>
                  <span className="text-stone-300">{g.categoria}</span>
                </div>
                {g.descripcion && <p className="text-xs text-stone-500 mt-0.5">{g.descripcion}</p>}
              </div>
              <div className="flex items-center gap-3">
                <span style={{ fontFamily: FONT_MONO }}>${money(g.monto)}</span>
                <button onClick={() => onDelete(g.id)} className="text-stone-600 active:text-red-400"><Trash2 size={14} /></button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function StockView({ productos = [], ventas = [], onAddProducto, onAddVenta, onDeleteVenta, onDeleteProducto }) {
  const [mostrarNuevoProd, setMostrarNuevoProd] = useState(false);
  const [prodSeleccionado, setProdSeleccionado] = useState(null);
  const [prodVentaId, setProdVentaId] = useState('');
  const [cantVenta, setCantVenta] = useState(1);
  const [medioPagoVenta, setMedioPagoVenta] = useState('efectivo');

  // Formulario nuevo producto
  const [nombreProd, setNombreProd] = useState('');
  const [stockProd, setStockProd] = useState('');
  const [costoProd, setCostoProd] = useState('');
  const [precioProd, setPrecioProd] = useState('');
  const [nicoProd, setNicoProd] = useState('');
  const [martinProd, setMartinProd] = useState('');

  // Estadísticas de inventario
  const totalInvertido = useMemo(() => {
    return (productos || []).reduce((acc, p) => acc + ((p.costo_unitario || 0) * (p.stock || 0)), 0);
  }, [productos]);

  const gananciaPotencial = useMemo(() => {
    return (productos || []).reduce((acc, p) => {
      const gananciaU = (p.precio_venta || 0) - (p.costo_unitario || 0);
      return acc + (gananciaU * (p.stock || 0));
    }, 0);
  }, [productos]);

  const stockBajoCount = useMemo(() => {
    return (productos || []).filter((p) => (p.stock || 0) <= 3).length;
  }, [productos]);

  // Selector de mes para ventas y rendiciones
  const [mesAnchor, setMesAnchor] = useState(new Date());
  const mesAnchorISO = `${mesAnchor.getFullYear()}-${pad2(mesAnchor.getMonth() + 1)}`;
  const cambiarMes = (n) => setMesAnchor(new Date(mesAnchor.getFullYear(), mesAnchor.getMonth() + n, 1));
  const esMesActual = mesAnchorISO === toISO(new Date()).slice(0, 7);

  const ventasMes = useMemo(() => {
    return (ventas || [])
      .filter((v) => (v.fecha || '').slice(0, 7) === mesAnchorISO)
      .sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
  }, [ventas, mesAnchorISO]);

  const totalNicoMes = useMemo(() => ventasMes.reduce((acc, v) => acc + (v.ganancia_nico || 0), 0), [ventasMes]);
  const totalMartinMes = useMemo(() => ventasMes.reduce((acc, v) => acc + (v.ganancia_martin || 0), 0), [ventasMes]);

  const submitProducto = (e) => {
    e.preventDefault();
    if (!nombreProd.trim()) return;

    onAddProducto({
      nombre: nombreProd.trim(),
      stock: parseInt(stockProd) || 0,
      costo_unitario: parseFloat(costoProd) || 0,
      precio_venta: parseFloat(precioProd) || 0,
      reparto_nico: parseFloat(nicoProd) || 0,
      reparto_martin: parseFloat(martinProd) || 0,
    });

    setNombreProd('');
    setStockProd('');
    setCostoProd('');
    setPrecioProd('');
    setNicoProd('');
    setMartinProd('');
    setMostrarNuevoProd(false);
  };

  const abrirVenta = (p) => {
    setProdSeleccionado(p);
    setProdVentaId(p.id);
    setCantVenta(1);
  };

  const submitVenta = (e) => {
    e.preventDefault();
    const prod = (productos || []).find((p) => p.id === prodVentaId);
    if (!prod) {
      alert('Seleccioná un producto válido');
      return;
    }

    const cantidad = parseInt(cantVenta) || 1;
    if (cantidad <= 0) return;

    if (prod.stock < cantidad) {
      alert(`¡No hay stock suficiente! Quedan solo ${prod.stock} unidades de ${prod.nombre}.`);
      return;
    }

    onAddVenta({
      producto_id: prod.id,
      producto_nombre: prod.nombre,
      cantidad,
      precio_total: (prod.precio_venta || 0) * cantidad,
      ganancia_nico: (prod.reparto_nico || 0) * cantidad,
      ganancia_martin: (prod.reparto_martin || 0) * cantidad,
      fecha: toISO(new Date()),
      medio_pago: medioPagoVenta,
      stock_actual: prod.stock
    });

    setProdSeleccionado(null);
  };

  return (
    <div className="space-y-6">
      {/* Tarjetas Superiores */}
      <div className="grid grid-cols-3 gap-2">
        <StatCard label="Inversión stock" valor={totalInvertido} />
        <StatCard label="Ganancia pot." valor={gananciaPotencial} tono="emerald" />
        <StatCard label="Stock bajo" valor={stockBajoCount} esNumero tono={stockBajoCount > 0 ? "red" : "amber"} />
      </div>

      {/* Rendición / Ganancias del mes con selector */}
      <section className="bg-stone-900 border border-stone-800 rounded-lg p-4 space-y-3">
        <div className="flex items-center justify-between">
          <button onClick={() => cambiarMes(-1)} className="p-1 text-stone-400 active:text-amber-400"><ChevronLeft size={18} /></button>
          <div className="text-center">
            <p className="text-xs uppercase tracking-widest text-stone-500" style={{ fontFamily: FONT_MONO }}>
              Rendición · {MESES[mesAnchor.getMonth()]} {mesAnchor.getFullYear()}
            </p>
            {!esMesActual && (
              <button onClick={() => setMesAnchor(new Date())} className="text-xs text-amber-400" style={{ fontFamily: FONT_MONO }}>volver a este mes</button>
            )}
          </div>
          <button onClick={() => cambiarMes(1)} className="p-1 text-stone-400 active:text-amber-400"><ChevronRight size={18} /></button>
        </div>
        <p className="text-xs text-amber-400 font-medium text-center" style={{ fontFamily: FONT_MONO }}>{ventasMes.length} ventas</p>
        <div className="grid grid-cols-2 gap-3">
          <div className="bg-stone-950 p-3 rounded-lg border border-stone-800">
            <p className="text-xs text-stone-400">Nico cobra:</p>
            <p className="text-lg font-bold text-amber-400" style={{ fontFamily: FONT_MONO }}>${money(totalNicoMes)}</p>
          </div>
          <div className="bg-stone-950 p-3 rounded-lg border border-stone-800">
            <p className="text-xs text-stone-400">Martín cobra:</p>
            <p className="text-lg font-bold text-amber-400" style={{ fontFamily: FONT_MONO }}>${money(totalMartinMes)}</p>
          </div>
        </div>
      </section>

      {/* Acciones principales */}
      <div className="flex gap-2">
        <button
          onClick={() => setMostrarNuevoProd(!mostrarNuevoProd)}
          className="flex-1 flex items-center justify-center gap-1.5 bg-stone-800 hover:bg-stone-700 active:bg-stone-600 text-stone-200 text-sm font-semibold rounded-lg py-2.5 transition-colors border border-stone-700"
        >
          <Plus size={16} /> {mostrarNuevoProd ? 'Ocultar formulario' : 'Nuevo Producto'}
        </button>
      </div>

      {/* Formulario Nuevo Producto */}
      {mostrarNuevoProd && (
        <form onSubmit={submitProducto} className="bg-stone-900 border border-stone-800 rounded-lg p-4 space-y-3">
          <p className="text-xs uppercase tracking-widest text-amber-400 font-semibold" style={{ fontFamily: FONT_MONO }}>+ Alta de Producto</p>
          <div className="space-y-2">
            <input
              type="text" placeholder="Nombre del producto" value={nombreProd}
              onChange={(e) => setNombreProd(e.target.value)} required
              className="w-full bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600 text-sm"
            />
            <div className="grid grid-cols-3 gap-2">
              <input
                type="number" placeholder="Stock" value={stockProd}
                onChange={(e) => setStockProd(e.target.value)} required
                className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600 text-sm" style={{ fontFamily: FONT_MONO }}
              />
              <input
                type="number" placeholder="Costo Unit. $" value={costoProd}
                onChange={(e) => setCostoProd(e.target.value)} required
                className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600 text-sm" style={{ fontFamily: FONT_MONO }}
              />
              <input
                type="number" placeholder="Precio Venta $" value={precioProd}
                onChange={(e) => setPrecioProd(e.target.value)} required
                className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600 text-sm" style={{ fontFamily: FONT_MONO }}
              />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <input
                type="number" placeholder="Reparto Nico $" value={nicoProd}
                onChange={(e) => setNicoProd(e.target.value)} required
                className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600 text-sm" style={{ fontFamily: FONT_MONO }}
              />
              <input
                type="number" placeholder="Reparto Martín $" value={martinProd}
                onChange={(e) => setMartinProd(e.target.value)} required
                className="bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100 placeholder-stone-600 text-sm" style={{ fontFamily: FONT_MONO }}
              />
            </div>
          </div>
          <button type="submit" className="w-full bg-amber-500 hover:bg-amber-400 text-stone-950 font-bold py-2 rounded text-sm">
            Guardar Producto
          </button>
        </form>
      )}

      {/* Modal/Formulario de Venta Rápida */}
      {prodSeleccionado && (
        <div className="fixed inset-0 bg-stone-950/80 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <form onSubmit={submitVenta} className="bg-stone-900 border border-stone-700 rounded-xl p-5 w-full max-w-sm space-y-4 shadow-2xl">
            <div className="flex items-center justify-between border-b border-stone-800 pb-3">
              <h3 className="font-semibold text-lg text-amber-400" style={{ fontFamily: FONT_DISPLAY }}>Registrar Venta</h3>
              <button type="button" onClick={() => setProdSeleccionado(null)} className="text-stone-400 hover:text-stone-200 text-sm">✕</button>
            </div>
            <div>
              <p className="text-stone-200 font-medium text-base">{prodSeleccionado.nombre}</p>
              <p className="text-xs text-stone-400 mt-0.5">Precio: ${money(prodSeleccionado.precio_venta)} · Stock disp: <strong className="text-amber-400">{prodSeleccionado.stock}</strong></p>
            </div>
            <div className="space-y-3">
              <div>
                <label className="text-xs text-stone-400 mb-1 block">Cantidad a vender:</label>
                <input
                  type="number" min="1" max={prodSeleccionado.stock} value={cantVenta}
                  onChange={(e) => setCantVenta(e.target.value)} required
                  className="w-full bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100" style={{ fontFamily: FONT_MONO }}
                />
              </div>
              <div>
                <label className="text-xs text-stone-400 mb-1 block">Medio de pago:</label>
                <select
                  value={medioPagoVenta} onChange={(e) => setMedioPagoVenta(e.target.value)}
                  className="w-full bg-stone-950 border border-stone-700 rounded px-3 py-2 text-stone-100"
                >
                  <option value="efectivo">Efectivo</option>
                  <option value="qr">QR / Transferencia</option>
                  <option value="tarjeta">Tarjeta</option>
                </select>
              </div>
            </div>
            <div className="bg-stone-950 p-3 rounded border border-stone-800 text-xs space-y-1" style={{ fontFamily: FONT_MONO }}>
              <div className="flex justify-between text-stone-400"><span>Total a cobrar:</span><span className="text-stone-100 font-bold">${money((prodSeleccionado.precio_venta || 0) * (parseInt(cantVenta) || 1))}</span></div>
              <div className="flex justify-between text-amber-400"><span>Reparto Nico:</span><span>${money((prodSeleccionado.reparto_nico || 0) * (parseInt(cantVenta) || 1))}</span></div>
              <div className="flex justify-between text-amber-400"><span>Reparto Martín:</span><span>${money((prodSeleccionado.reparto_martin || 0) * (parseInt(cantVenta) || 1))}</span></div>
            </div>
            <div className="flex gap-2">
              <button type="button" onClick={() => setProdSeleccionado(null)} className="flex-1 bg-stone-800 hover:bg-stone-700 text-stone-300 py-2 rounded text-sm font-semibold">
                Cancelar
              </button>
              <button type="submit" className="flex-1 bg-amber-500 hover:bg-amber-400 text-stone-950 py-2 rounded text-sm font-bold">
                Confirmar Venta
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Listado de Inventario */}
      <section className="space-y-3">
        <h3 className="text-sm uppercase tracking-widest text-stone-500" style={{ fontFamily: FONT_MONO }}>Inventario de Productos</h3>
        <div className="grid grid-cols-1 gap-2">
          {(productos || []).map((p) => {
            const stockBajo = (p.stock || 0) <= 3;
            return (
              <div key={p.id} className="bg-stone-900 border border-stone-800 rounded-lg p-3 flex items-center justify-between">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-stone-200" style={{ fontFamily: FONT_DISPLAY }}>{p.nombre}</span>
                    {stockBajo && (
                      <span className="bg-red-950/80 border border-red-800 text-red-400 text-[10px] uppercase tracking-wider px-1.5 py-0.5 rounded font-bold">
                        Bajo
                      </span>
                    )}
                  </div>
                  <div className="flex items-center gap-3 text-xs text-stone-400" style={{ fontFamily: FONT_MONO }}>
                    <span>Stock: <strong className={stockBajo ? "text-red-400" : "text-emerald-400"}>{p.stock}</strong> u.</span>
                    <span>·</span>
                    <span>Venta: ${money(p.precio_venta)}</span>
                    <span>·</span>
                    <span>N: ${money(p.reparto_nico)} / M: ${money(p.reparto_martin)}</span>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => abrirVenta(p)}
                    disabled={p.stock <= 0}
                    className={`flex items-center gap-1 text-xs font-bold px-3 py-2 rounded-lg transition-colors ${
                      p.stock > 0 
                        ? 'bg-amber-500 hover:bg-amber-400 text-stone-950 active:scale-95' 
                        : 'bg-stone-800 text-stone-600 cursor-not-allowed'
                    }`}
                  >
                    <ShoppingCart size={14} /> Vender
                  </button>
                  <button
                    onClick={() => onDeleteProducto(p.id)}
                    title="Borrar producto"
                    className="text-stone-600 active:text-red-400 p-2"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </section>

      {/* Historial de Ventas del mes seleccionado */}
      <section className="space-y-3">
        <h3 className="text-sm uppercase tracking-widest text-stone-500" style={{ fontFamily: FONT_MONO }}>
          Ventas de {MESES[mesAnchor.getMonth()]}
        </h3>
        {ventasMes.length === 0 ? (
          <p className="text-stone-500 text-sm">No hay ventas registradas en {MESES[mesAnchor.getMonth()]}.</p>
        ) : (
          <div className="bg-stone-900 border border-stone-800 rounded-lg divide-y divide-stone-800">
            {ventasMes.map((v) => (
              <div key={v.id} className="flex items-center justify-between px-4 py-2.5 text-sm">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-stone-200">{v.producto_nombre}</span>
                    <span className="text-xs text-amber-400 font-bold" style={{ fontFamily: FONT_MONO }}>x{v.cantidad}</span>
                  </div>
                  <div className="flex items-center gap-2 text-xs text-stone-500 mt-0.5" style={{ fontFamily: FONT_MONO }}>
                    <span>{v.fecha ? formatShort(fromISO(v.fecha)) : ''}</span>
                    <span>·</span>
                    <span className="capitalize">{v.medio_pago || 'efectivo'}</span>
                    <span>·</span>
                    <span>Nico: ${money(v.ganancia_nico)}</span>
                    <span>·</span>
                    <span>Martín: ${money(v.ganancia_martin)}</span>
                  </div>
                </div>
                <div className="flex items-center gap-3">
                  <span className="font-bold text-stone-100" style={{ fontFamily: FONT_MONO }}>${money(v.precio_total)}</span>
                  <button onClick={() => onDeleteVenta(v)} title="Anular venta y reponer stock" className="text-stone-600 active:text-red-400">
                    <Trash2 size={14} />
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </section>
    </div>
  );
}

function TrendCard({ label, valor, variacion }) {
  const tono = variacion > 0 ? 'text-emerald-400' : variacion < 0 ? 'text-red-400' : 'text-stone-500';
  return (
    <div className="bg-stone-900 border border-stone-800 rounded-lg p-3">
      <p className="text-xs text-stone-500 mb-1">{label}</p>
      <p className="text-lg font-semibold text-stone-100" style={{ fontFamily: FONT_MONO }}>${money(valor)}</p>
      {variacion !== null && variacion !== undefined && (
        <p className={`text-xs flex items-center gap-1 mt-1 ${tono}`} style={{ fontFamily: FONT_MONO }}>
          {variacion > 0 ? <TrendingUp size={12} /> : variacion < 0 ? <TrendingDown size={12} /> : null}
          {variacion > 0 ? '+' : ''}{Math.round(variacion)}% vs. período anterior
        </p>
      )}
    </div>
  );
}

function calcularVariacion(arr) {
  return arr.map((item, i) => {
    if (i === 0) return { ...item, variacion: null };
    const prev = arr[i - 1].facturado;
    const variacion = prev > 0 ? ((item.facturado - prev) / prev) * 100 : (item.facturado > 0 ? 100 : 0);
    return { ...item, variacion };
  });
}

function BarraTendencia({ datos, maxValor }) {
  return (
    <div className="flex items-end gap-2 h-32">
      {datos.map((d) => (
        <div key={d.key} className="flex-1 flex flex-col items-center justify-end gap-1 h-full">
          {d.variacion !== null && d.variacion !== undefined ? (
            <span className={`text-xs flex items-center gap-0.5 ${d.variacion > 0 ? 'text-emerald-400' : d.variacion < 0 ? 'text-red-400' : 'text-stone-500'}`}>
              {d.variacion > 0 ? <TrendingUp size={10} /> : d.variacion < 0 ? <TrendingDown size={10} /> : null}
              {d.variacion > 0 ? '+' : ''}{Math.round(d.variacion)}%
            </span>
          ) : <span className="text-xs text-stone-600">—</span>}
          <div
            className={`w-full rounded-t ${d.variacion > 0 ? 'bg-emerald-500' : d.variacion < 0 ? 'bg-red-500' : 'bg-stone-600'}`}
            style={{ height: `${Math.max(4, (d.facturado / maxValor) * 100)}%` }}
          />
          <span className="text-xs text-stone-500 capitalize">{d.label}</span>
        </div>
      ))}
    </div>
  );
}

function ReportesView({ registros = [], gastos = [], pct }) {
  const hoy = new Date();

  const mensual = useMemo(() => {
    const meses = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
      const key = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
      const cortes = (registros || []).filter((r) => r.tipo === 'corte' && (r.fecha || '').slice(0, 7) === key);
      const ropas = (registros || []).filter((r) => r.tipo === 'ropa' && (r.fecha || '').slice(0, 7) === key);
      const totalCortes = cortes.reduce((s, r) => s + r.monto, 0);
      const totalRopa = ropas.reduce((s, r) => s + r.monto, 0);
      const facturado = totalCortes + totalRopa;
      const paraBarberia = totalCortes * (1 - pct / 100) + totalRopa;
      const gastosMes = (gastos || []).filter((g) => (g.fecha || '').slice(0, 7) === key).reduce((s, g) => s + g.monto, 0);
      meses.push({ key, label: `${MESES[d.getMonth()].slice(0, 3)} '${String(d.getFullYear()).slice(2)}`, facturado, gananciaNeta: paraBarberia - gastosMes });
    }
    return calcularVariacion(meses);
  }, [registros, gastos, pct]);

  const semanal = useMemo(() => {
    const semanas = [];
    const lunesActual = mondayOf(hoy);
    for (let i = 7; i >= 0; i--) {
      const lunes = addDays(lunesActual, -7 * i);
      const sabado = addDays(lunes, 5);
      const isoL = toISO(lunes), isoS = toISO(sabado);
      const cortes = (registros || []).filter((r) => r.tipo === 'corte' && (r.fecha || '') >= isoL && (r.fecha || '') <= isoS);
      const ropas = (registros || []).filter((r) => r.tipo === 'ropa' && (r.fecha || '') >= isoL && (r.fecha || '') <= isoS);
      const facturado = cortes.reduce((s, r) => s + r.monto, 0) + ropas.reduce((s, r) => s + r.monto, 0);
      semanas.push({ key: isoL, label: formatShort(lunes), facturado });
    }
    return calcularVariacion(semanas);
  }, [registros]);

  const maxMensual = Math.max(1, ...mensual.map((m) => m.facturado));
  const maxSemanal = Math.max(1, ...semanal.map((s) => s.facturado));
  const ultimoMes = mensual[mensual.length - 1];
  const ultimaSemana = semanal[semanal.length - 1];

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 gap-3">
        <TrendCard label="Este mes" valor={ultimoMes.facturado} variacion={ultimoMes.variacion} />
        <TrendCard label="Esta semana" valor={ultimaSemana.facturado} variacion={ultimaSemana.variacion} />
      </div>

      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Facturación por mes (últimos 6)</h3>
        <div className="bg-stone-900 border border-stone-800 rounded-lg p-4">
          <BarraTendencia datos={mensual} maxValor={maxMensual} />
        </div>
        <div className="mt-2 divide-y divide-stone-800">
          {mensual.slice().reverse().map((m) => (
            <div key={m.key} className="flex items-center justify-between text-sm py-1.5">
              <span className="capitalize text-stone-300">{m.label}</span>
              <div className="flex items-center gap-3">
                <span style={{ fontFamily: FONT_MONO }}>${money(m.facturado)}</span>
                <span className="text-stone-500 text-xs" style={{ fontFamily: FONT_MONO }}>ganancia ${money(m.gananciaNeta)}</span>
              </div>
            </div>
          ))}
        </div>
      </section>

      <section>
        <h3 className="text-sm uppercase tracking-widest text-stone-500 mb-2" style={{ fontFamily: FONT_MONO }}>Facturación por semana (últimas 8)</h3>
        <div className="bg-stone-900 border border-stone-800 rounded-lg p-4">
          <BarraTendencia datos={semanal} maxValor={maxSemanal} />
        </div>
      </section>
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);
  const [syncError, setSyncError] = useState(false);
  const [config, setConfig] = useState({ barberos: ['Nico', 'Martín', 'Gonza'], comisionPct: 50 });
  const [registros, setRegistros] = useState([]);
  const [gastos, setGastos] = useState([]);
  const [productos, setProductos] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [tab, setTab] = useState('hoy');
  const [dateSel, setDateSel] = useState(toISO(new Date()));
  const [monthAnchor, setMonthAnchor] = useState(new Date());
  const [weekAnchor, setWeekAnchor] = useState(mondayOf(new Date()));

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setLoading(false);
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      setLoading(false);
    });

    return () => {
      if (subscription) subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (!session) return;

    async function cargarDatos() {
      try {
        const { data: regData } = await supabase.from('registros').select('*');
        if (regData) setRegistros(regData);

        const { data: gastData } = await supabase.from('gastos').select('*');
        if (gastData) setGastos(gastData);

        const { data: confData } = await supabase.from('config').select('*').eq('id', 1).single();
        if (confData) {
          setConfig({ barberos: confData.barberos || [], comisionPct: confData.comision_pct ?? 50 });
        }

        const { data: prodData } = await supabase.from('productos').select('*').order('nombre');
        if (prodData) setProductos(prodData);

        const { data: ventData } = await supabase.from('ventas').select('*').order('created_at', { ascending: false });
        if (ventData) setVentas(ventData);

      } catch (err) {
        console.error('Error cargando datos:', err);
      }
    }

    cargarDatos();
  }, [session]);

  const persistConfig = async (nuevoConfig) => {
    setConfig(nuevoConfig);
    const { error } = await supabase
      .from('config')
      .upsert({ id: 1, barberos: nuevoConfig.barberos, comision_pct: nuevoConfig.comisionPct });

    if (error) {
      console.error('Error guardando config:', error);
      setSyncError(true);
    }
  };

  const addRegistro = async (entry) => {
    const nuevo = { id: uid(), ...entry };
    const { error } = await supabase.from('registros').insert([nuevo]);
    if (error) {
      console.error('Error guardando registro:', error);
      setSyncError(true);
      return;
    }
    setRegistros((prev) => [...prev, nuevo]);
    setSyncError(false);
  };

  const deleteRegistro = async (id) => {
    const { error } = await supabase.from('registros').delete().eq('id', id);
    if (error) {
      console.error('Error eliminando registro:', error);
      setSyncError(true);
      return;
    }
    setRegistros((prev) => prev.filter((r) => r.id !== id));
  };

  const addGasto = async (entry) => {
    const nuevo = { id: uid(), ...entry };
    const { error } = await supabase.from('gastos').insert([nuevo]);
    if (error) {
      console.error('Error guardando gasto:', error);
      setSyncError(true);
      return;
    }
    setGastos((prev) => [...prev, nuevo]);
    setSyncError(false);
  };

  const deleteGasto = async (id) => {
    const { error } = await supabase.from('gastos').delete().eq('id', id);
    if (error) {
      console.error('Error eliminando gasto:', error);
      setSyncError(true);
      return;
    }
    setGastos((prev) => prev.filter((g) => g.id !== id));
  };

  // Funciones de Stock y Ventas
  const addProducto = async (entry) => {
    const nuevo = { id: uid(), ...entry };
    const { error } = await supabase.from('productos').insert([nuevo]);
    if (error) {
      console.error('Error guardando producto:', error);
      setSyncError(true);
      return;
    }
    setProductos((prev) => [...prev, nuevo].sort((a, b) => (a.nombre || '').localeCompare(b.nombre || '')));
    setSyncError(false);
  };

  const deleteProducto = async (id) => {
    if (!window.confirm('¿Seguro que querés borrar este producto? No se borran las ventas ya registradas, solo desaparece del inventario.')) {
      return;
    }
    const { error } = await supabase.from('productos').delete().eq('id', id);
    if (error) {
      console.error('Error eliminando producto:', error);
      setSyncError(true);
      return;
    }
    setProductos((prev) => prev.filter((p) => p.id !== id));
  };

  const addVenta = async (ventaData) => {
    const { stock_actual, ...restoVenta } = ventaData;
    const nuevaVenta = { id: uid(), ...restoVenta };
    const nuevoStock = stock_actual - ventaData.cantidad;

    // 1. Registrar la venta
    const { error: errorVenta } = await supabase.from('ventas').insert([nuevaVenta]);
    if (errorVenta) {
      console.error('Error al guardar venta:', errorVenta);
      setSyncError(true);
      return;
    }

    // 2. Descontar el stock en productos
    const { error: errorStock } = await supabase
      .from('productos')
      .update({ stock: nuevoStock })
      .eq('id', ventaData.producto_id);

    if (errorStock) {
      console.error('Error actualizando stock:', errorStock);
      setSyncError(true);
    }

    // 3. Actualizar estados locales
    setVentas((prev) => [nuevaVenta, ...prev]);
    setProductos((prev) =>
      prev.map((p) => (p.id === ventaData.producto_id ? { ...p, stock: nuevoStock } : p))
    );
    setSyncError(false);
  };

  const deleteVenta = async (venta) => {
    if (!window.confirm(`¿Seguro que querés anular esta venta de ${venta.producto_nombre}? Se repondrán ${venta.cantidad} unidades al stock.`)) {
      return;
    }

    // 1. Eliminar la venta
    const { error: errorDelete } = await supabase.from('ventas').delete().eq('id', venta.id);
    if (errorDelete) {
      console.error('Error eliminando venta:', errorDelete);
      setSyncError(true);
      return;
    }

    // 2. Reponer el stock
    const prod = productos.find((p) => p.id === venta.producto_id);
    const nuevoStock = (prod?.stock || 0) + (venta.cantidad || 0);

    if (venta.producto_id) {
      await supabase.from('productos').update({ stock: nuevoStock }).eq('id', venta.producto_id);
    }

    // 3. Actualizar estados locales
    setVentas((prev) => prev.filter((v) => v.id !== venta.id));
    if (venta.producto_id) {
      setProductos((prev) =>
        prev.map((p) => (p.id === venta.producto_id ? { ...p, stock: nuevoStock } : p))
      );
    }
  };
  const exportarBackup = () => {
    const backup = {
      exportado_el: new Date().toISOString(),
      config,
      registros,
      gastos,
      productos,
      ventas,
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `backup-barberia-${toISO(new Date())}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const generarHoja = (workbook, nombreHoja, columnas, datos) => {
    const hoja = workbook.addWorksheet(nombreHoja);
    hoja.columns = columnas.map((c) => ({ header: c.header, key: c.key, width: c.width || 16 }));

    hoja.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF292524' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    (datos || []).forEach((fila) => {
      const filaFormateada = {};
      columnas.forEach((c) => { filaFormateada[c.key] = fila[c.key]; });
      hoja.addRow(filaFormateada);
    });

    columnas.forEach((c, i) => {
      if (c.moneda) hoja.getColumn(i + 1).numFmt = '$ #,##0.00';
    });

    const colsMoneda = columnas.filter((c) => c.moneda);
    if (colsMoneda.length > 0 && datos && datos.length > 0) {
      const filaTotales = {};
      columnas.forEach((c) => { filaTotales[c.key] = ''; });
      filaTotales[columnas[0].key] = 'TOTAL';
      colsMoneda.forEach((c) => {
        filaTotales[c.key] = datos.reduce((s, f) => s + (Number(f[c.key]) || 0), 0);
      });
      const rowTotal = hoja.addRow(filaTotales);
      rowTotal.eachCell((cell) => {
        cell.font = { bold: true };
        cell.border = { top: { style: 'thin' } };
      });
    }

    if (columnas.length > 0) {
      hoja.autoFilter = { from: 'A1', to: { row: 1, column: columnas.length } };
    }
  };

  const exportarExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Libro de Caja - Barbería';
    workbook.created = new Date();

    generarHoja(workbook, 'Registros', [
      { header: 'Fecha', key: 'fecha', width: 12 },
      { header: 'Barbero', key: 'barbero', width: 14 },
      { header: 'Tipo', key: 'tipo', width: 12 },
      { header: 'Método', key: 'metodo', width: 12 },
      { header: 'Hora', key: 'hora', width: 10 },
      { header: 'Monto', key: 'monto', width: 14, moneda: true },
    ], registros);

    generarHoja(workbook, 'Gastos', [
      { header: 'Fecha', key: 'fecha', width: 12 },
      { header: 'Categoría', key: 'categoria', width: 18 },
      { header: 'Descripción', key: 'descripcion', width: 26 },
      { header: 'Monto', key: 'monto', width: 14, moneda: true },
    ], gastos);

    generarHoja(workbook, 'Productos', [
      { header: 'Nombre', key: 'nombre', width: 22 },
      { header: 'Stock', key: 'stock', width: 10 },
      { header: 'Costo Unit.', key: 'costo_unitario', width: 14, moneda: true },
      { header: 'Precio Venta', key: 'precio_venta', width: 14, moneda: true },
      { header: 'Reparto Nico', key: 'reparto_nico', width: 14, moneda: true },
      { header: 'Reparto Martín', key: 'reparto_martin', width: 16, moneda: true },
    ], productos);

    generarHoja(workbook, 'Ventas', [
      { header: 'Fecha', key: 'fecha', width: 12 },
      { header: 'Producto', key: 'producto_nombre', width: 22 },
      { header: 'Cantidad', key: 'cantidad', width: 10 },
      { header: 'Medio de pago', key: 'medio_pago', width: 16 },
      { header: 'Total', key: 'precio_total', width: 14, moneda: true },
      { header: 'Ganancia Nico', key: 'ganancia_nico', width: 14, moneda: true },
      { header: 'Ganancia Martín', key: 'ganancia_martin', width: 16, moneda: true },
    ], ventas);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `libro-caja-${toISO(new Date())}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  const irAFecha = (iso) => { setDateSel(iso); setTab('hoy'); };

  if (loading) {
    return (
      <div className="min-h-screen bg-stone-950 flex items-center justify-center text-stone-400" style={{ fontFamily: FONT_BODY }}>
        <Loader2 className="animate-spin mr-2" size={20} /> Cargando...
      </div>
    );
  }

  if (!session) return <Auth />;

  return (
    <div className="min-h-screen bg-stone-950 text-stone-100 pb-24" style={{ fontFamily: FONT_BODY }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');`}</style>
      <BarraBarbero />
      <header className="px-4 pt-5 pb-3 border-b border-stone-800">
        <div className="flex items-baseline justify-between">
          <div>
            <h1 className="text-2xl font-semibold tracking-tight" style={{ fontFamily: FONT_DISPLAY }}>Libro de Caja</h1>
            <span className="text-xs text-stone-500 uppercase tracking-widest" style={{ fontFamily: FONT_MONO }}>Barbería</span>
          </div>

                   <div className="flex items-center gap-2">
            <button
              onClick={exportarExcel}
              className="bg-stone-800 hover:bg-stone-700 text-stone-200 text-xs font-bold px-3 py-2 rounded-lg transition-colors border border-stone-700"
            >
              Excel
            </button>
            <button
              onClick={exportarBackup}
              className="bg-stone-800 hover:bg-stone-700 text-stone-200 text-xs font-bold px-3 py-2 rounded-lg transition-colors border border-stone-700"
            >
              Backup
            </button>
            <button 
              onClick={() => supabase.auth.signOut()}
              className="bg-red-600 hover:bg-red-700 text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors"
            >
              Cerrar Sesión
            </button>
          </div>
        </div>
        {syncError && <p className="mt-2 text-xs text-red-400">No se pudo guardar el último cambio. Revisá tu conexión.</p>}
      </header>  

      <main className="px-4 py-4">
        {tab === 'hoy' && (
          <HoyView dateSel={dateSel} setDateSel={setDateSel} barberos={config.barberos} pct={config.comisionPct} registros={registros} addRegistro={addRegistro} deleteRegistro={deleteRegistro} />
        )}
        {tab === 'calendario' && (
          <CalendarioView monthAnchor={monthAnchor} setMonthAnchor={setMonthAnchor} registros={registros} onSelectDay={irAFecha} dateSel={dateSel} />
        )}
        {tab === 'cierres' && (
          <CierresView weekAnchor={weekAnchor} setWeekAnchor={setWeekAnchor} barberos={config.barberos} pct={config.comisionPct} registros={registros} gastos={gastos} />
        )}
        {tab === 'gastos' && (
          <GastosView gastos={gastos} onAdd={addGasto} onDelete={deleteGasto} />
        )}
        {tab === 'stock' && (
          <StockView productos={productos} ventas={ventas} onAddProducto={addProducto} onAddVenta={addVenta} onDeleteVenta={deleteVenta} onDeleteProducto={deleteProducto} />
        )}
        {tab === 'reportes' && (
          <ReportesView registros={registros} gastos={gastos} pct={config.comisionPct} />
        )}
        {tab === 'barberos' && (
          <BarberosView config={config} onSave={persistConfig} />
        )}
      </main>

      <nav className="fixed bottom-0 left-0 right-0 bg-stone-900 border-t border-stone-800 flex z-50 overflow-x-auto">
        {[
          { key: 'hoy', label: 'Hoy', Icon: ListChecks },
          { key: 'calendario', label: 'Calendario', Icon: CalendarDays },
          { key: 'cierres', label: 'Cierres', Icon: Scissors },
          { key: 'gastos', label: 'Gastos', Icon: Wallet },
          { key: 'stock', label: 'Stock', Icon: Package },
          { key: 'reportes', label: 'Reportes', Icon: TrendingUp },
          { key: 'barberos', label: 'Barberos', Icon: Users },
        ].map(({ key, label, Icon }) => (
          <button key={key} onClick={() => setTab(key)} className={`flex-1 min-w-[55px] flex flex-col items-center gap-1 py-2.5 ${tab === key ? 'text-amber-400' : 'text-stone-500'}`}>
            <Icon size={18} />
            <span className="text-[11px]" style={{ fontFamily: FONT_MONO }}>{label}</span>
          </button>
        ))}
      </nav>
    </div>
  );
}