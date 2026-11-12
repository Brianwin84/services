/* Mis Vehículos — lógica compartida entre la app y el service worker */
(function (g) {
  const normalPat = p => String(p || '').replace(/\s+/g, '').toUpperCase();
  const usaHoras = t => t === 'Embarcación' || t === 'Aeronave';

  function cleanFecha(str) {
    if (!str || String(str).trim() === '' || /ERROR|Invalid|S\/F|undefined/.test(str)) return '';
    return String(str).substring(0, 10);
  }
  function parseFecha(str) {
    const s = cleanFecha(str); if (!s) return null;
    const d = new Date(s + 'T12:00:00'); return isNaN(d.getTime()) ? null : d;
  }
  function hoy0() { const d = new Date(); d.setHours(12, 0, 0, 0); return d; }
  function diffDias(d, base) { return Math.round((d - (base || hoy0())) / 86400000); }
  function num(v) { if (v === '' || v === null || v === undefined) return null; const n = Number(v); return isNaN(n) ? null : n; }
  function rowN(s) { const n = Number(s.idRow); return isNaN(n) ? Infinity : n; }
  function partes(tipo) { return String(tipo || '').split('+').map(s => s.trim()).filter(Boolean); }

  /** km/hs actual: máximo entre el registrado y el de cualquier servicio */
  function kmActual(data, pat) {
    const p = normalPat(pat);
    const v = (data.vehiculos || []).find(x => normalPat(x.patente) === p);
    let km = v ? (num(v.kmActual) || 0) : 0;
    (data.servicios || []).forEach(s => { if (normalPat(s.patente) === p) { const k = num(s.km); if (k && k > km) km = k; } });
    return km;
  }

  /** Un servicio está "activo" si es el más reciente para al menos uno de sus ítems */
  function activos(data) {
    const ult = {};
    (data.servicios || []).forEach(s => {
      partes(s.tipo).forEach(c => {
        const k = normalPat(s.patente) + '|' + c.toLowerCase();
        const u = ult[k];
        const fs = cleanFecha(s.fecha), fu = u && cleanFecha(u.fecha);
        if (!u || fs > fu || (fs === fu && rowN(s) > rowN(u))) ult[k] = s;
      });
    });
    // Un service queda "reemplazado" si hay uno más nuevo del mismo vehículo
    // que repite al menos uno de sus ítems (ej: un nuevo cambio de aceite).
    const set = new Set();
    const porVeh = {};
    (data.servicios || []).forEach(s => { const p = normalPat(s.patente); (porVeh[p] = porVeh[p] || []).push(s); });
    Object.values(porVeh).forEach(list => {
      list.sort((a, b) => (cleanFecha(b.fecha) || '').localeCompare(cleanFecha(a.fecha) || '') || rowN(b) - rowN(a));
      const vistos = new Set();
      list.forEach(s => {
        const its = partes(s.tipo).map(x => x.toLowerCase());
        if (!its.some(i => vistos.has(i))) set.add(String(s.idRow));
        its.forEach(i => vistos.add(i));
      });
    });
    return { set, ult };
  }

  /** Estado de un servicio: level = venc | prox | ok | none */
  function estado(data, s, cfg) {
    const v = (data.vehiculos || []).find(x => normalPat(x.patente) === normalPat(s.patente));
    const hs = usaHoras(v && v.tipoVeh);
    const unidad = hs ? 'hs' : 'km';
    const margen = hs ? (cfg.margenHs || 20) : (cfg.margenKm || 1000);
    const dAviso = cfg.diasAviso || 30;
    const pf = parseFecha(s.proxFecha);
    const dias = pf ? diffDias(pf) : null;
    const pk = num(s.proxKm);
    const act = kmActual(data, s.patente);
    const rest = (pk && pk > 0 && act) ? pk - act : null;

    let level = 'none';
    if (dias !== null || rest !== null) level = 'ok';
    if ((dias !== null && dias <= dAviso) || (rest !== null && rest <= margen)) level = 'prox';
    if ((dias !== null && dias < 0) || (rest !== null && rest <= 0)) level = 'venc';

    // progreso consumido (0..1) — el más crítico entre fecha y km
    let prog = null;
    const f0 = parseFecha(s.fecha);
    if (f0 && pf && pf > f0) prog = Math.max(0, Math.min(1, (hoy0() - f0) / (pf - f0)));
    const k0 = num(s.km);
    if (pk && k0 !== null && pk > k0 && act) {
      const pk2 = Math.max(0, Math.min(1, (act - k0) / (pk - k0)));
      prog = prog === null ? pk2 : Math.max(prog, pk2);
    }
    return { level, dias, rest, unidad, prog, kmAct: act };
  }

  /** Alertas: servicios activos vencidos o próximos, ordenados por urgencia */
  function alertas(data, cfg) {
    const { set } = activos(data);
    const out = [];
    (data.servicios || []).forEach(s => {
      if (!set.has(String(s.idRow))) return;
      const e = estado(data, s, cfg);
      if (e.level === 'venc' || e.level === 'prox') out.push({ s, e });
    });
    const score = x => (x.e.level === 'venc' ? -100000 : 0) +
      Math.min(x.e.dias === null ? 99999 : x.e.dias, x.e.rest === null ? 99999 : x.e.rest / 50);
    return out.sort((a, b) => score(a) - score(b));
  }

  function textoEstado(e) {
    const t = [];
    if (e.dias !== null) t.push(e.dias < 0 ? `vencido hace ${-e.dias} d` : e.dias === 0 ? 'vence hoy' : `faltan ${e.dias} d`);
    if (e.rest !== null) t.push(e.rest <= 0 ? `excedido ${Math.abs(e.rest).toLocaleString('es-AR')} ${e.unidad}` : `faltan ${e.rest.toLocaleString('es-AR')} ${e.unidad}`);
    return t.join(' · ');
  }

  g.VTK = { normalPat, usaHoras, cleanFecha, parseFecha, diffDias, hoy0, num, partes, kmActual, activos, estado, alertas, textoEstado };
})(typeof self !== 'undefined' ? self : window);
