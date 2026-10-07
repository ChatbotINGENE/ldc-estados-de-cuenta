/* Lector del Excel de estados de cuenta de Lomas de California.
   Es la misma lógica de ../01-tablero-cobranza/cobranza.py, en JavaScript,
   para que todo corra en el navegador de quien usa la página (el Excel no
   se sube a ningún lado). Si cambia una, hay que cambiar la otra. */
(function () {
  var SUPUESTOS = { diasGracia: 30 };
  var GRUPO_APARTE = ["LDCG10", "LDCG11", "LDCG12", "LDCH1", "LDCH5", "LDCH20"];
  var NOMBRE_GRUPO = "Grupo Iraheta";
  var PRIMA = /prima|reserv|complemento|dep[oó]sito/i;
  var AJENO = /chapoda|reparaci|casa|mantenim|limpieza/i;
  // Avisos que no cambian el saldo: el PDF sí sale.
  var NO_BLOQUEA = [/la fecha está escrita como texto/, /La prima en la ficha es igual al precio/, /La cuota anotada/];
  var MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

  var PAC = "America/Los_Angeles", MTN = "America/Denver", CEN = "America/Chicago", EST = "America/New_York";
  var AREAS = {};
  [
    ["213 310 323 424 562 626 747 818", "Los Ángeles", PAC], ["657 714 949", "Orange County", PAC],
    ["909 951 760 442", "Sur de California", PAC], ["408 415 510 650 669 628", "Área de San Francisco", PAC],
    ["209 559 916 805 661", "California", PAC], ["206 253 425 360", "Washington (estado)", PAC],
    ["702 725", "Las Vegas", PAC], ["303 720 801 385", "Colorado / Utah", MTN],
    ["281 346 713 832", "Houston", CEN], ["214 469 972 817 682", "Dallas", CEN], ["210 512 737", "Texas", CEN],
    ["312 773 630 847 872", "Chicago", CEN], ["504 985", "Luisiana", CEN],
    ["240 301 410 443 667", "Maryland", EST], ["202", "Washington D. C.", EST], ["571 703", "Virginia", EST],
    ["212 332 347 646 718 917 929", "Nueva York", EST], ["516 631 934", "Long Island", EST],
    ["201 551 732 848 862 908 973", "Nueva Jersey", EST], ["617 781 857 978", "Massachusetts", EST],
    ["305 786 954 754 561", "Sur de Florida", EST], ["321 407 689", "Orlando", EST],
    ["404 470 678 770", "Atlanta", EST], ["704 980 919 984", "Carolina del Norte", EST], ["215 267 445", "Filadelfia", EST]
  ].forEach(function (a) { a[0].split(" ").forEach(function (c) { AREAS[c] = [a[1], a[2]]; }); });

  function ubicarTelefono(tel) {
    if (tel == null || tel === "") return ["Sin teléfono", null];
    var d = String(tel).replace(/\D/g, "");
    if (d.indexOf("503") === 0 && d.length === 11) return ["El Salvador", "America/El_Salvador"];
    if (d.length === 8 && "267".indexOf(d[0]) >= 0) return ["El Salvador", "America/El_Salvador"];
    if (d.length === 11 && d[0] === "1") d = d.slice(1);
    if (d.length === 10) return AREAS[d.slice(0, 3)] || ["EE. UU. (código " + d.slice(0, 3) + ")", null];
    return ["Sin identificar", null];
  }

  // ── Fechas: todas en UTC ──
  function iso(d) { return d.toISOString().slice(0, 10); }
  function aFecha(v) {
    if (typeof v === "number" && v > 20000 && v < 80000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    if (typeof v === "string") {
      var m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) return new Date(Date.UTC(+m[3], +m[1] - 1, +m[2]));
      m = v.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    }
    return null;
  }
  function sumarMeses(d, n) {
    var m = d.getUTCMonth() + n, y = d.getUTCFullYear() + Math.floor(m / 12);
    m = ((m % 12) + 12) % 12;
    var ultimo = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
    return new Date(Date.UTC(y, m, Math.min(d.getUTCDate(), ultimo)));
  }
  function dias(a, b) { return Math.round((a - b) / 86400000); }
  function primeroDeMes(d) { return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1)); }
  function claveMes(d) { return iso(d).slice(0, 7); }
  function fechaCorta(d) { if (!d) return "—"; var p = iso(d).split("-"); return p[2] + "/" + p[1] + "/" + p[0]; }
  function fechaLarga(d) { return d.getUTCDate() + " de " + MESES[d.getUTCMonth()] + " de " + d.getUTCFullYear(); }
  function dinero(x, dec) {
    if (x == null) return "—";
    var d = dec == null ? 2 : dec;
    return "$" + Number(x).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
  }
  function f0(x) { return Number(x).toLocaleString("en-US", { maximumFractionDigits: 0 }); }
  function num(v) { return typeof v === "number" ? v : null; }

  function cuotaFija(monto, tasaAnual, meses) {
    var r = tasaAnual / 12;
    return r ? monto * r / (1 - Math.pow(1 + r, -meses)) : monto / meses;
  }
  // Lotes que cubre un contrato, según el nombre de la hoja (LDCG14G15 → G 14 y 15).
  function lotesDeHoja(codigo) {
    var m = codigo.toUpperCase().match(/^LDC([A-Z])/);
    return { poligono: m ? m[1] : "", lotes: (codigo.match(/\d+/g) || []).map(Number) };
  }

  function leerHoja(ws, nombreHoja, corte) {
    var codigo = nombreHoja.split(/\s+/)[0];
    function val(r, c) {
      var cell = ws[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })];
      if (!cell || cell.v === "" || cell.v == null) return null;
      return cell.v;
    }
    var celdas = {}, filaTabla = null;
    for (var r = 1; r <= 60; r++) for (var c = 1; c <= 20; c++) {
      var v = val(r, c);
      if (typeof v === "string") {
        var k = v.trim().toLowerCase();
        if (!(k in celdas)) celdas[k] = { r: r, c: c };
        if (v.trim() === "No." && filaTabla === null) filaTabla = r;
      }
    }
    function junto(et) { var p = celdas[et]; return p ? val(p.r, p.c + 1) : null; }
    var cliente = junto("cliente");
    if (!cliente || filaTabla === null) return null;

    var tel = junto("tel.");
    var enc = celdas["precio de contado"];
    var fd = enc ? enc.r + 2 : 18, col = enc ? enc.c : 5;
    var precio = num(val(fd, col)) || 0;
    var primaEnc = num(val(fd, col + 1));
    var textos = [fd - 2, fd - 1, fd].map(function (x) { var t = val(x, col + 2); return t == null ? "" : String(t); }).join(" ");
    var g18 = val(fd, col + 2), loteTxt = val(fd, col - 2), poligono = val(fd, col - 3), area = val(fd, col - 1);
    var revisar = [];

    var plan;
    if (/contado/i.test(textos)) plan = "contado";
    else if (/15%|inter[eé]s 15|anual/i.test(textos) && !/sin inter/i.test(textos)) plan = "credito";
    else plan = "sin_intereses";

    var cuota = null, plazoM = null, m;
    m = textos.match(/cuota\s*\$\s*([\d,]+(?:\.\d+)?)/i);
    if (m) cuota = parseFloat(m[1].replace(/,/g, ""));
    else if (plan === "credito" && typeof g18 === "number") cuota = g18;
    m = textos.match(/(\d+)\s*(a[ñn]os?|anos?)/i);
    if (m) plazoM = parseInt(m[1], 10) * 12;

    var loteStr = typeof loteTxt === "number" ? String(Math.trunc(loteTxt)) : String(loteTxt == null ? "" : loteTxt);
    var esperado = codigo.replace(/^LDC[A-Z]0?/i, "");
    var enHoja = (codigo.match(/\d+/g) || []).map(Number).sort().join(",");
    var enFicha = (loteStr.match(/\d+/g) || []).map(Number).sort().join(",");
    if (enHoja && enFicha && enHoja !== enFicha) revisar.push("La hoja se llama " + codigo + " pero la ficha dice lote " + loteStr + ".");

    var pagos = [], saldoHoja = null, saldoPrev = null;
    for (var rr = filaTabla + 1; rr < filaTabla + 40; rr++) {
      var fRaw = val(rr, 3), desc = val(rr, 4), pago = num(val(rr, 5)), capital = num(val(rr, 6)), interes = num(val(rr, 7)), saldo = num(val(rr, 8));
      if (fRaw === null && pago === null) continue;
      var fecha = aFecha(fRaw);
      if (typeof fRaw === "string") revisar.push("Fila " + rr + ": la fecha está escrita como texto (" + fRaw + ").");
      if (!fecha || pago === null) { revisar.push("Fila " + rr + ": falta la fecha o el monto."); continue; }
      if (plan !== "credito" && capital !== null && Math.abs(capital - pago) > 0.5)
        revisar.push("Fila " + rr + " (" + desc + "): el pago es $" + f0(pago) + " pero el capital abonado dice $" + f0(capital) + ".");
      if (desc && AJENO.test(String(desc)))
        revisar.push("Fila " + rr + ": «" + desc + "» por $" + f0(pago) + " rebaja el saldo del lote. ¿Es un pago del lote?");
      if (saldoPrev !== null && saldo !== null && capital !== null && Math.abs(saldoPrev - capital - saldo) > 1)
        revisar.push("Fila " + rr + ": el saldo no cuadra con el saldo anterior menos el capital.");
      if (saldo !== null) { saldoPrev = saldo; saldoHoja = saldo; }
      pagos.push({ fecha: fecha, desc: String(desc == null ? "" : desc).trim(), pago: pago, capital: capital, interes: interes, saldo: saldo, esPrima: !!(desc && PRIMA.test(String(desc))) });
    }
    pagos.sort(function (a, b) { return a.fecha - b.fecha; });
    if (saldoHoja === null) saldoHoja = num(junto("saldo actual")) || precio;
    var saldoFicha = num(junto("saldo actual"));
    if (saldoFicha !== null && Math.abs(saldoFicha - saldoHoja) > 1)
      revisar.push("El «Saldo actual» del encabezado ($" + f0(saldoFicha) + ") no coincide con el último saldo de la tabla ($" + f0(saldoHoja) + ").");
    if (primaEnc && precio && Math.abs(primaEnc - precio) < 1 && plan !== "contado")
      revisar.push("La prima en la ficha es igual al precio ($" + f0(precio) + "). Probablemente es un error de captura.");

    var pc = pagos.filter(function (p) { return p.fecha <= corte; });
    var primas = pc.filter(function (p) { return p.esPrima; });
    var inicio = null;
    if (primas.length) {
      inicio = primas[primas.length - 1].fecha;
      pc.forEach(function (p) {
        if (!p.esPrima && p.fecha < primas[0].fecha)
          revisar.push("El abono del " + fechaCorta(p.fecha) + " por $" + f0(p.pago) + " es anterior a la prima (" + fechaCorta(primas[0].fecha) + "). ¿El año está mal escrito?");
      });
    } else if (pc.length) inicio = pc[0].fecha;
    var ultimo = pc.length ? pc[pc.length - 1].fecha : null;
    var pagado = pc.reduce(function (s, p) { return s + p.pago; }, 0);
    var primaPagada = primas.reduce(function (s, p) { return s + p.pago; }, 0);
    var abonos = pagado - primaPagada;
    var lugar = ubicarTelefono(tel);

    var c2 = {
      codigo: codigo, cliente: String(cliente).trim(), telefono: tel == null ? "" : String(tel).trim(),
      dui: junto("dui") || junto("dpi"), poligono: String(poligono == null ? "" : poligono).trim(), lote: loteStr || esperado,
      area: area, precio: precio, plan: plan, cuota: cuota, plazoM: plazoM, saldo: Math.max(0, saldoHoja), pagado: pagado,
      inicio: inicio, ultimoPago: ultimo, diasSinPagar: ultimo ? dias(corte, ultimo) : null,
      grupo: GRUPO_APARTE.indexOf(codigo) >= 0 ? NOMBRE_GRUPO : "", lugar: lugar[0], tz: lugar[1],
      revisar: revisar, pagos: pagos
    };

    var atraso = 0, diasAtraso = 0, proximo = null, vence = null, esperadoMes = 0, flujo = {}, estado, etiqueta = null;
    var base = primeroDeMes(corte);
    if (c2.saldo <= 1) estado = "pagado";
    else if (plan === "contado") {
      estado = "revisar";
      revisar.push("Está marcado como compra de contado pero tiene saldo pendiente de $" + f0(c2.saldo) + ".");
    } else if (plan === "credito" && cuota && inicio) {
      var vencidas = 0;
      while (sumarMeses(inicio, vencidas + 1) <= corte) vencidas++;
      atraso = Math.max(0, vencidas * cuota - abonos);
      var cubiertas = Math.floor(abonos / cuota + 1e-9);
      if (cubiertas < vencidas) diasAtraso = dias(corte, sumarMeses(inicio, cubiertas + 1));
      proximo = sumarMeses(inicio, Math.max(vencidas, cubiertas) + 1);
      esperadoMes = cuota;
      for (var k = 1; k <= 12; k++) flujo[claveMes(sumarMeses(base, k))] = cuota;
      estado = "al_dia";
    } else if (plan === "sin_intereses" && plazoM && inicio) {
      vence = sumarMeses(inicio, plazoM);
      var financiado = Math.max(0, precio - primaPagada);
      var transcurrido = Math.max(0, dias(corte, inicio)), totalDias = Math.max(1, dias(vence, inicio));
      atraso = Math.max(0, financiado * Math.min(1, transcurrido / totalDias) - abonos);
      var porDia = financiado ? financiado / totalDias : 0;
      diasAtraso = porDia ? Math.floor(atraso / porDia) : 0;
      if (corte >= vence && c2.saldo > 1) diasAtraso = Math.max(diasAtraso, dias(corte, vence));
      esperadoMes = financiado / plazoM;
      var restante = Math.max(0, c2.saldo - atraso), meses = [];
      for (var j = 1; j <= 12; j++) { var mes = sumarMeses(base, j); if (mes <= primeroDeMes(vence)) meses.push(claveMes(mes)); }
      meses.forEach(function (mk) { flujo[mk] = restante / meses.length; });
      estado = "al_dia";
    } else {
      estado = "revisar";
      revisar.push("No se pudo leer el plan de pagos (plazo o cuota).");
    }

    if (estado === "al_dia" && atraso >= 1) {
      if (plan === "credito") estado = diasAtraso > SUPUESTOS.diasGracia ? "atraso" : "atraso_leve";
      else {
        var sinAbonar = c2.diasSinPagar || 0;
        if (vence && corte >= vence) { estado = "atraso"; etiqueta = "Plazo vencido"; }
        else if (sinAbonar > 60) { estado = "atraso"; etiqueta = "Sin abonar " + sinAbonar + " días"; }
        else { estado = "atraso_leve"; etiqueta = "Va lento"; }
      }
    }

    // La cuota del crédito debe salir del precio menos la prima, a 10 años al 15% (cuota fija).
    if (plan === "credito" && cuota && plazoM && precio && primaEnc != null && primaEnc < precio) {
      var debida = cuotaFija(precio - primaEnc, 0.15, plazoM);
      if (Math.abs(debida - cuota) > 1)
        revisar.push("La cuota anotada ($" + cuota.toFixed(2) + ") no corresponde al precio y la prima de la ficha: debería ser $" + debida.toFixed(2) + ".");
    }

    c2.estado = estado; c2.etiqueta = etiqueta; c2.atraso = atraso; c2.diasAtraso = diasAtraso;
    c2.proximo = proximo; c2.vence = vence; c2.esperadoMes = esperadoMes; c2.flujo = flujo;
    c2.recibidoMes = pc.filter(function (p) { return claveMes(p.fecha) === claveMes(corte); }).reduce(function (s, p) { return s + p.pago; }, 0);
    c2.bloqueos = revisar.filter(function (x) { return !NO_BLOQUEA.some(function (re) { return re.test(x); }); });
    return c2;
  }

  function leerLibro(libro, corte) {
    var contratos = [], libres = 0;
    libro.SheetNames.forEach(function (n) {
      if (!/^LDC/i.test(n)) return;
      var c = leerHoja(libro.Sheets[n], n, corte);
      if (c) contratos.push(c); else libres++;
    });
    return { contratos: contratos, libres: libres };
  }

  // ── Hoja «Gastos»: A fecha · B descripción · C monto · D a nombre de quién ──
  // Gastos que no son del proyecto (se muestran aparte y no cuentan como inversión).
  var FUERA = /pr[eé]stamo a bazar|compra veh[ií]culo|pet scan/i;
  var CATEGORIAS = [
    ["Comisiones de venta", /comisi/i],
    ["Viajes", /viaje/i],
    ["Calles y terracería", /calle|terracer|balastad|baden|compactaci/i],
    ["Topografía y mojones", /amojon|mojon|levantamiento|replante|topograf|trazo|linea divisoria/i],
    ["Planos y trámites (CNR)", /cnr|plano|registral|escritura|digitalizaci/i],
    ["Chapoda y limpieza", /chapod|limpieza|mozos/i],
    ["Impuestos y contabilidad", /impuesto|contabilidad/i],
    ["Diseño y rotulación", /dise[ñn]o|rotulaci/i],
    ["Préstamos a personas", /pr[eé]stamo/i],
    ["Promoción y redes", /redes|publicidad/i]
  ];

  function leerGastos(libro, corte) {
    var ws = libro.Sheets["Gastos"];
    if (!ws) return null;
    var rango = XLSX.utils.decode_range(ws["!ref"] || "A1:D1");
    function val(r, c) { var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })]; return cell && cell.v !== "" ? cell.v : null; }
    var gastos = [], avisos = [];
    for (var r = rango.s.r; r <= rango.e.r; r++) {
      var fRaw = val(r, 0), desc = val(r, 1), monto = val(r, 2), quien = val(r, 3);
      if (fRaw === null && desc === null && monto === null) continue;
      var fila = r + 1, fecha = aFecha(fRaw), d = String(desc == null ? "" : desc).trim();
      if (!fecha) { avisos.push({ fila: fila, texto: "Fila " + fila + " («" + (d || "sin descripción") + "»): falta la fecha o no se entiende.", frena: false }); }
      if (typeof monto !== "number") { avisos.push({ fila: fila, texto: "Fila " + fila + " («" + (d || "sin descripción") + "»): falta el monto.", frena: false }); continue; }
      if (!d) avisos.push({ fila: fila, texto: "Fila " + fila + ": gasto de $" + f0(monto) + " sin descripción.", frena: false });
      if (fecha && fecha > corte) avisos.push({ fila: fila, texto: "Fila " + fila + " («" + d + "»): la fecha " + fechaCorta(fecha) + " es posterior a la fecha del Excel.", frena: false });
      var cat = "Otros";
      for (var i = 0; i < CATEGORIAS.length; i++) if (CATEGORIAS[i][1].test(d)) { cat = CATEGORIAS[i][0]; break; }
      var fuera = FUERA.test(d);
      if (cat === "Préstamos a personas" && !fuera)
        avisos.push({ fila: fila, texto: "Fila " + fila + ": «" + d + "» por $" + f0(monto) + ". ¿Es un gasto del proyecto o un préstamo que se va a devolver?", frena: false });
      gastos.push({ fila: fila, fecha: fecha, desc: d, monto: monto, quien: String(quien == null ? "" : quien).trim() || "Sin nombre", categoria: fuera ? "No es del proyecto" : cat, fuera: fuera });
    }
    // Posibles duplicados: misma fecha, misma descripción y mismo monto.
    var grupos = {};
    gastos.forEach(function (g) {
      if (!g.fecha) return;
      var k = iso(g.fecha) + "|" + g.desc.toLowerCase() + "|" + g.monto;
      (grupos[k] = grupos[k] || []).push(g);
    });
    // Si el mismo gasto se repite en muchas fechas (p. ej. viajes de ida y vuelta), va en un solo aviso.
    var porGasto = {};
    Object.keys(grupos).forEach(function (k) {
      var gs = grupos[k]; if (gs.length < 2) return;
      var kk = gs[0].desc.toLowerCase() + "|" + gs[0].monto;
      (porGasto[kk] = porGasto[kk] || []).push(gs);
    });
    Object.keys(porGasto).forEach(function (kk) {
      var sets = porGasto[kk], g0 = sets[0][0];
      var extra = sets.reduce(function (s, gs) { return s + (gs.length - 1) * g0.monto; }, 0);
      if (sets.length === 1) {
        var gs = sets[0];
        avisos.push({ fila: g0.fila, frena: false, texto: "«" + g0.desc + "» por $" + f0(g0.monto) + " aparece " + gs.length + " veces el " + fechaCorta(g0.fecha) +
          " (filas " + gs.map(function (g) { return g.fila; }).join(", ") + "). ¿Son gastos distintos o está repetido?" });
      } else {
        avisos.push({ fila: g0.fila, frena: false, texto: "«" + g0.desc + "» por $" + f0(g0.monto) + " aparece más de una vez el mismo día en " + sets.length +
          " fechas (" + sets.map(function (gs) { return fechaCorta(gs[0].fecha); }).join(", ") + "). Si son ida y vuelta o dos personas, está bien; si no, hay $" + f0(extra) + " repetidos." });
      }
    });
    avisos.sort(function (a, b) { return a.fila - b.fila; });
    return { gastos: gastos, avisos: avisos };
  }

  // ── Hoja «Lista de precios», cruzada con los estados de cuenta ──
  function leerInventario(libro, contratos) {
    var ws = libro.Sheets["Lista de precios"];
    if (!ws) return null;
    var rango = XLSX.utils.decode_range(ws["!ref"] || "A1:I1");
    function val(r, c) { var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })]; return cell && cell.v !== "" ? cell.v : null; }
    var lotes = [], avisos = [], porClave = {};
    for (var r = rango.s.r + 1; r <= rango.e.r; r++) {
      var pol = val(r, 0), n = val(r, 1);
      if (pol == null || n == null) continue;
      var bloqueado = /bloque/i.test(String(val(r, 2) || "")) || /bloque/i.test(String(n));
      var l = {
        fila: r + 1, poligono: String(pol).trim().toUpperCase(), lote: typeof n === "number" ? Math.trunc(n) : parseInt(n, 10),
        area: num(val(r, 2)), lista: num(val(r, 3)), venta: num(val(r, 4)), prima: num(val(r, 5)),
        vendedor: val(r, 6) ? String(val(r, 6)).trim() : "", plan: val(r, 7) ? String(val(r, 7)).trim() : "",
        cancelado: /pagad|cancel/i.test(String(val(r, 8) || "")), bloqueado: bloqueado, contrato: null
      };
      if (isNaN(l.lote)) continue;
      lotes.push(l); porClave[l.poligono + l.lote] = l;
    }
    // Cruce con los contratos (el estado de cuenta manda: si hay hoja con cliente, está vendido).
    var difArea = {};
    contratos.forEach(function (c) {
      var h = lotesDeHoja(c.codigo), ls = [];
      h.lotes.forEach(function (n) {
        var l = porClave[h.poligono + n];
        if (!l) { avisos.push({ texto: "La hoja " + c.codigo + " (" + c.cliente + ") es del lote " + h.poligono + "-" + n + ", que no está en la lista de precios." }); return; }
        l.contrato = c; ls.push(l);
        if (l.venta == null) avisos.push({ texto: "Lote " + l.poligono + "-" + l.lote + ": tiene estado de cuenta (" + c.cliente + ") pero en la lista de precios no tiene precio de venta." });
        if (l.bloqueado) avisos.push({ texto: "Lote " + l.poligono + "-" + l.lote + ": la lista lo marca BLOQUEADO pero tiene estado de cuenta (" + c.cliente + ")." });
        if (l.area != null && typeof c.area === "number" && h.lotes.length === 1 && Math.abs(l.area - c.area) > 0.5) {
          var ka = l.area.toFixed(2) + "|" + c.area.toFixed(2);
          (difArea[ka] = difArea[ka] || []).push(l.poligono + "-" + l.lote);
        }
        if (l.cancelado && c.saldo > 1) avisos.push({ texto: "Lote " + l.poligono + "-" + l.lote + ": la lista dice «Pagado» pero el estado de cuenta tiene saldo de $" + f0(c.saldo) + "." });
        if (!l.cancelado && c.saldo <= 1) avisos.push({ texto: "Lote " + l.poligono + "-" + l.lote + ": ya está pagado según el estado de cuenta, pero la lista no lo marca como «Pagado»." });
      });
      var ventaLista = ls.reduce(function (s, l) { return s + (l.venta || 0); }, 0);
      if (ls.length && ventaLista && Math.abs(ventaLista - c.precio) > 1)
        avisos.push({ texto: "Lote " + h.poligono + "-" + h.lotes.join("/") + ": el precio de venta es $" + f0(ventaLista) + " en la lista y $" + f0(c.precio) + " en el estado de cuenta." });
    });
    // Las diferencias de área iguales van juntas: el estado de cuenta (lo que ve el cliente) y la lista no coinciden.
    Object.keys(difArea).forEach(function (ka) {
      var v = ka.split("|"), ls2 = difArea[ka];
      avisos.unshift({ texto: (ls2.length > 1 ? "Lotes " + ls2.join(", ") + ": el área" : "Lote " + ls2[0] + ": el área") + " es " + v[0] + " v² en la lista de precios y " + v[1] +
        " v² en " + (ls2.length > 1 ? "sus estados de cuenta" : "su estado de cuenta") + ". El estado de cuenta es lo que recibe el cliente: ¿cuál es la correcta?" });
    });
    lotes.forEach(function (l) {
      if (l.venta != null && !l.contrato && !l.bloqueado)
        avisos.push({ texto: "Lote " + l.poligono + "-" + l.lote + ": la lista dice que se vendió en $" + f0(l.venta) + (l.vendedor ? " (" + l.vendedor + ")" : "") + ", pero no hay estado de cuenta con cliente." });
      l.estado = l.bloqueado ? "bloqueado" : (l.contrato || l.venta != null) ? "vendido" : "disponible";
    });
    return { lotes: lotes, avisos: avisos };
  }

  window.LDC = {
    leerLibro: leerLibro, leerGastos: leerGastos, leerInventario: leerInventario, cuotaFija: cuotaFija, SUPUESTOS: SUPUESTOS, NOMBRE_GRUPO: NOMBRE_GRUPO, MESES: MESES,
    iso: iso, aFecha: aFecha, sumarMeses: sumarMeses, fechaCorta: fechaCorta, fechaLarga: fechaLarga, dinero: dinero
  };
})();
