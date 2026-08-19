/* ==========================================================================
   app.js
   Application controller: state, DOM rendering, event wiring.
   Depends on Calculator (js/calculator.js) and PptxGenerator (js/pptx-generator.js).
   ========================================================================== */

(function () {
  'use strict';

  if (typeof window.ResizeObserver === 'undefined') {
    window.ResizeObserver = class {
      constructor(cb) { this.cb = cb; }
      observe() {}
      unobserve() {}
      disconnect() {}
    };
  }

  const C = window.Calculator;
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));

  /* ---------------------------- State ---------------------------- */

  let state = deepClone(C.SKYLUXE_PRESET);
  let model = C.computeModel(state);
  let currentTab = 'planner';
  let currentSlide = 0;

  function deepClone(o) { return JSON.parse(JSON.stringify(o)); }

  function recalc() {
    model = C.computeModel(state);
    updateSidebarSummary();
    if (currentTab === 'deck') renderDeck();
    if (currentTab === 'export') renderExportSummary();
  }

  function setState(patch) {
    Object.assign(state, patch);
    recalc();
  }

  function setDeep(key, patch) {
    state[key] = Object.assign({}, state[key], patch);
    recalc();
  }

  /* ---------------------------- Small DOM helpers ---------------------------- */

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(k => {
      if (k === 'class') node.className = attrs[k];
      else if (k === 'html') node.innerHTML = attrs[k];
      else if (k.indexOf('on') === 0 && typeof attrs[k] === 'function') node.addEventListener(k.slice(2), attrs[k]);
      else node.setAttribute(k, attrs[k]);
    });
    (children || []).forEach(c => {
      if (c === null || c === undefined) return;
      node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return node;
  }

  function numberInput(value, onChange, opts) {
    opts = opts || {};
    const input = el('input', { type: 'number', class: 'num-input', step: opts.step || 1, value: value });
    input.addEventListener('input', e => onChange(e.target.value === '' ? 0 : Number(e.target.value)));
    return input;
  }

  function monthMatrix(labelText, months, values, onChange) {
    const wrap = el('div', { class: 'mb-3' });
    wrap.appendChild(el('span', { class: 'field-label' }, [labelText]));
    const grid = el('div', { class: 'grid gap-2', style: `grid-template-columns:repeat(${months.length},minmax(0,1fr))` });
    months.forEach((m, i) => {
      const col = el('div', {}, [
        el('div', { class: 'text-[10px] text-slate-400 mb-1 truncate' }, [m]),
        numberInput(values[i] || 0, v => {
          const next = values.slice(); next[i] = v; onChange(next);
        }),
      ]);
      grid.appendChild(col);
    });
    wrap.appendChild(grid);
    return wrap;
  }

  function field(labelText, inputNode, hint) {
    const wrap = el('label', { class: 'block mb-3' });
    wrap.appendChild(el('span', { class: 'field-label' }, [labelText]));
    wrap.appendChild(inputNode);
    if (hint) wrap.appendChild(el('span', { class: 'block text-[11px] text-slate-400 mt-1' }, [hint]));
    return wrap;
  }

  function sectionCard(title, iconHtml, contentNode, defaultOpen) {
    const card = el('div', { class: 'card overflow-hidden mb-4' });
    const header = el('button', { class: 'section-toggle', type: 'button' }, [
      el('div', { class: 'flex items-center gap-3' }, [
        el('div', { class: 'icon-badge', html: iconHtml }),
        el('span', { class: 'font-semibold text-slate-800 text-sm tracking-wide' }, [title]),
      ]),
      el('span', { class: 'chevron', html: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M9 18l6-6-6-6"/></svg>' }),
    ]);
    const body = el('div', { class: 'px-5 pb-5 pt-1' }, [contentNode]);
    body.style.display = defaultOpen === false ? 'none' : 'block';
    header.querySelector('.chevron').style.transform = defaultOpen === false ? 'rotate(0deg)' : 'rotate(90deg)';
    header.addEventListener('click', () => {
      const open = body.style.display !== 'none';
      body.style.display = open ? 'none' : 'block';
      header.querySelector('.chevron').style.transform = open ? 'rotate(0deg)' : 'rotate(90deg)';
    });
    card.appendChild(header);
    card.appendChild(body);
    return card;
  }

  /* ---------------------------- Planner Tab ---------------------------- */

  function onDurationChange(val) {
    state = C.resizeStateToDuration(state, val, null);
    recalc();
    renderPlanner();
  }

  function renderPlanner() {
    const root = $('#tab-planner');
    root.innerHTML = '';

    const heading = el('div', { class: 'mb-6' }, [
      el('h2', { class: 'font-display text-2xl font-bold text-slate-900' }, ['Strategy Planner']),
      el('p', { class: 'text-sm text-slate-500 mt-1' }, ['Configure sourcing targets, site-visit split, digital benchmarks and offline spends. Every field flows live into the deck and export.']),
      el('div', { class: 'ruler mt-3 w-40' }),
    ]);
    root.appendChild(heading);

    const presetsRow = el('div', { class: 'flex flex-wrap gap-2 mb-6' }, [
      el('button', { class: 'preset-btn preset-dark', type: 'button', onclick: () => { state = deepClone(C.SKYLUXE_PRESET); recalc(); renderPlanner(); } }, ['Load Skyluxe 3-Month Plan (43 Bookings)']),
      el('button', { class: 'preset-btn preset-gold', type: 'button', onclick: () => { state = C.fastTrackPreset(); recalc(); renderPlanner(); } }, ['Load Fast-Track Plan (20 Bookings / 80 Visits, 6-6-8 Split)']),
      el('button', { class: 'preset-btn preset-outline', type: 'button', onclick: () => { state = C.makeState({ projectName: 'NEW PROJECT', bookingsPS: [0,0,0], bookingsDirect: [0,0,0], bookingsCP: [0,0,0], visitsPS: [0,0,0], visitsDirect: [0,0,0], visitsCP: [0,0,0] }); recalc(); renderPlanner(); } }, ['Reset to Custom']),
    ]);
    root.appendChild(presetsRow);

    /* Basic parameters */
    const basicGrid = el('div', { class: 'grid grid-cols-2 gap-x-4' });
    basicGrid.appendChild(field('Project Name', (() => {
      const i = el('input', { class: 'text-input', value: state.projectName });
      i.addEventListener('input', e => setState({ projectName: e.target.value }));
      return i;
    })()));
    basicGrid.appendChild(field('Duration (Months)', (() => {
      const i = el('input', { type: 'number', min: 1, max: 12, class: 'num-input', value: state.duration });
      i.addEventListener('change', e => onDurationChange(e.target.value));
      return i;
    })()));
    basicGrid.appendChild(field('Month Names (comma separated)', (() => {
      const i = el('input', { class: 'text-input', value: state.months.join(', ') });
      i.addEventListener('input', e => setState({ months: e.target.value.split(',').map(s => s.trim()) }));
      return i;
    })(), 'Order matches duration'));
    basicGrid.appendChild(field('Average Agreement Value / Ticket Size', numberInput(state.avgTicketSize, v => setState({ avgTicketSize: v }), { step: 100000 })));
    basicGrid.appendChild(field('Booking-to-Visit Conversion Benchmark', numberInput((state.conversionRate * 100).toFixed(2), v => setState({ conversionRate: Number(v) / 100 }), { step: 0.5 }), 'Seeds the auto-fill helper below — not applied automatically'));
    basicGrid.appendChild(field('CP Brokerage %', numberInput((state.cpBrokeragePercent * 100).toFixed(2), v => setState({ cpBrokeragePercent: Number(v) / 100 }), { step: 0.5 })));
    root.appendChild(sectionCard('Basic Parameters', ICONS.building, basicGrid));

    /* Bookings */
    const bookingsWrap = el('div');
    bookingsWrap.appendChild(monthMatrix('Presales / Digital Bookings', model.months, state.bookingsPS, v => setState({ bookingsPS: v })));
    bookingsWrap.appendChild(monthMatrix('Direct / Walk-in Bookings', model.months, state.bookingsDirect, v => setState({ bookingsDirect: v })));
    bookingsWrap.appendChild(monthMatrix('Channel Partner Bookings', model.months, state.bookingsCP, v => setState({ bookingsCP: v })));
    const bookingsTotals = el('div', { class: 'flex gap-6 text-xs font-mono text-slate-500 mt-2 border-t pt-3', id: 'bookings-totals' });
    bookingsWrap.appendChild(bookingsTotals);
    root.appendChild(sectionCard('Target Bookings & Monthly Split', ICONS.target, bookingsWrap));

    /* Site Visits — directly editable, per source */
    const visitsWrap = el('div');
    visitsWrap.appendChild(el('p', { class: 'text-xs text-slate-500 mb-3' }, ['Set your site-visit targets directly — e.g. 80 visits across 3 months — split by source. The implied conversion % is calculated from what you enter here, it doesn\'t drive the numbers.']));
    const autoFillRow = el('div', { class: 'flex justify-end mb-2' });
    const autoFillBtn = el('button', { type: 'button', class: 'autofill-btn' }, [`Auto-fill from ${(state.conversionRate * 100).toFixed(1)}% Benchmark`]);
    autoFillBtn.addEventListener('click', () => {
      setState({
        visitsPS: C.deriveVisits(state.bookingsPS, state.conversionRate),
        visitsDirect: C.deriveVisits(state.bookingsDirect, state.conversionRate),
        visitsCP: C.deriveVisits(state.bookingsCP, state.conversionRate),
      });
      renderPlanner();
    });
    autoFillRow.appendChild(autoFillBtn);
    visitsWrap.appendChild(autoFillRow);
    visitsWrap.appendChild(monthMatrix('Presales / Digital Visits', model.months, state.visitsPS, v => setState({ visitsPS: v })));
    visitsWrap.appendChild(monthMatrix('Direct / Walk-in Visits', model.months, state.visitsDirect, v => setState({ visitsDirect: v })));
    visitsWrap.appendChild(monthMatrix('Channel Partner Visits', model.months, state.visitsCP, v => setState({ visitsCP: v })));
    const visitsWarning = el('div', { class: 'warning-banner', id: 'visits-warning' });
    visitsWrap.appendChild(visitsWarning);
    const visitsTotals = el('div', { class: 'grid grid-cols-4 gap-3 text-xs font-mono text-slate-500 mt-2 border-t pt-3', id: 'visits-totals' });
    visitsWrap.appendChild(visitsTotals);
    root.appendChild(sectionCard('Site Visits Target & Source Split', ICONS.mapPin, visitsWrap));

    /* Digital benchmark CPLs & platform ratios */
    const digitalWrap = el('div');
    digitalWrap.appendChild(monthMatrix('Cost Per Lead (CPL) by Month (₹)', model.months, state.cplMonthly, v => setState({ cplMonthly: v })));
    digitalWrap.appendChild(el('div', { class: 'text-[11px] font-semibold text-slate-500 mb-2 uppercase tracking-wide mt-4' }, ['Platform Weight of Lead-Gen Budget & Benchmark CPL']));
    const platformRows = el('div', { class: 'space-y-3' });
    [['meta', 'Meta / Facebook'], ['pmax', 'Google Pmax + Demand Gen'], ['search', 'Google Search'], ['experimental', 'Experimental / LinkedIn / Native'], ['nri', 'NRI Campaigns — USA/GCC']].forEach(([key, label]) => {
      const row = el('div', { class: 'grid grid-cols-12 items-center gap-3' });
      row.appendChild(el('span', { class: 'col-span-4 text-xs text-slate-600' }, [label]));
      const range = el('input', { type: 'range', min: 0, max: 100, step: 0.5, class: 'col-span-4', value: state.platformPct[key] * 100 });
      range.addEventListener('input', e => setDeep('platformPct', { [key]: Number(e.target.value) / 100 }));
      const pctLabel = el('span', { class: 'col-span-1 text-xs font-mono text-amber-600 text-right' }, [(state.platformPct[key] * 100).toFixed(1) + '%']);
      range.addEventListener('input', () => { pctLabel.textContent = (state.platformPct[key] * 100).toFixed(1) + '%'; });
      row.appendChild(range);
      row.appendChild(pctLabel);
      const cplCol = el('div', { class: 'col-span-3' }, [numberInput(state.platformCPL[key], v => setDeep('platformCPL', { [key]: v }))]);
      row.appendChild(cplCol);
      platformRows.appendChild(row);
    });
    digitalWrap.appendChild(platformRows);
    root.appendChild(sectionCard('Digital Benchmark CPLs & Platform Ratios', ICONS.sliders, digitalWrap));

    /* OOH, production & CP incentive settings */
    const oohGrid = el('div', { class: 'grid grid-cols-2 gap-x-4' });
    oohGrid.appendChild(field('Social Media Boosting', numberInput(state.socialMediaBoosting, v => setState({ socialMediaBoosting: v }), { step: 5000 })));
    oohGrid.appendChild(field('Drone & Construction Shoot', numberInput(state.droneShoot, v => setState({ droneShoot: v }), { step: 5000 })));
    oohGrid.appendChild(field('UGC Videos (10 Videos)', numberInput(state.ugcVideos, v => setState({ ugcVideos: v }), { step: 5000 })));
    oohGrid.appendChild(field('Developer Byte & Testimonial Shoot', numberInput(state.developerByteShoot, v => setState({ developerByteShoot: v }), { step: 5000 })));
    oohGrid.appendChild(field('Chatbot + WhatsApp API', numberInput(state.chatbotWhatsapp, v => setState({ chatbotWhatsapp: v }), { step: 5000 })));
    oohGrid.appendChild(field('Society & Mall Activations', numberInput(state.societyMallActivations, v => setState({ societyMallActivations: v }), { step: 10000 })));
    oohGrid.appendChild(field('Kiosks / Wayboards', numberInput(state.kiosks, v => setState({ kiosks: v }), { step: 5000 })));
    oohGrid.appendChild(field('Newspaper Leaflet Inserts', numberInput(state.newspaperLeaflet, v => setState({ newspaperLeaflet: v }), { step: 5000 })));
    oohGrid.appendChild(field('Channel Partner Gathering', numberInput(state.cpGathering, v => setState({ cpGathering: v }), { step: 5000 })));
    oohGrid.appendChild(field('CP High Tea', numberInput(state.cpHighTea, v => setState({ cpHighTea: v }), { step: 5000 })));
    oohGrid.appendChild(field('CP Incentive per Visit', numberInput(state.cpVisitIncentive, v => setState({ cpVisitIncentive: v }), { step: 50 })));
    oohGrid.appendChild(field('CP Incentive per Booking', numberInput(state.cpBookingIncentive, v => setState({ cpBookingIncentive: v }), { step: 1000 })));
    const oohWrap = el('div', {}, [oohGrid]);
    oohWrap.appendChild(el('div', { class: 'text-[11px] font-semibold text-slate-500 mb-2 uppercase tracking-wide mt-2' }, ['Hoardings / OOH']));
    const hoardGrid = el('div', { class: 'grid grid-cols-3 gap-4' }, [
      field('Count', numberInput(state.hoardings.count, v => setDeep('hoardings', { count: v }))),
      field('Months', numberInput(state.hoardings.months, v => setDeep('hoardings', { months: v }))),
      field('Rate / Month', numberInput(state.hoardings.rate, v => setDeep('hoardings', { rate: v }), { step: 5000 })),
    ]);
    oohWrap.appendChild(hoardGrid);
    oohWrap.appendChild(el('div', { class: 'text-xs text-slate-500 mt-1', id: 'hoarding-total' }));
    root.appendChild(sectionCard('OOH, Production & CP Incentive Settings', ICONS.megaphone, oohWrap));

    /* CP activation targets */
    const cpWrap = el('div');
    cpWrap.appendChild(monthMatrix('CAT A CP Activation Targets', model.months, state.catA, v => setState({ catA: v })));
    cpWrap.appendChild(monthMatrix('CAT B CP Activation Targets', model.months, state.catB, v => setState({ catB: v })));
    cpWrap.appendChild(monthMatrix('CP Digital Reach', model.months, state.cpDigitalReach, v => setState({ cpDigitalReach: v })));
    root.appendChild(sectionCard('Channel Partner Activation Targets', ICONS.users, cpWrap, false));

    /* Historical / reporting data */
    root.appendChild(sectionCard('Historical Reporting Data (Slides 2–5)', ICONS.calendar, renderHistoricalEditors(), false));

    updateComputedBits();
  }

  function updateComputedBits() {
    const bt = $('#bookings-totals');
    if (bt) {
      bt.innerHTML = '';
      bt.appendChild(el('span', {}, ['PS: ', el('b', { class: 'text-slate-800' }, [String(model.bookingsPS_total)])]));
      bt.appendChild(el('span', {}, ['Direct: ', el('b', { class: 'text-slate-800' }, [String(model.bookingsDirect_total)])]));
      bt.appendChild(el('span', {}, ['CP: ', el('b', { class: 'text-slate-800' }, [String(model.bookingsCP_total)])]));
      bt.appendChild(el('span', {}, ['Total: ', el('b', { class: 'text-amber-600' }, [String(model.totalBookings)])]));
    }
    const vw = $('#visits-warning');
    if (vw) {
      const bad = model.visitsPS_total < model.bookingsPS_total || model.visitsDirect_total < model.bookingsDirect_total || model.visitsCP_total < model.bookingsCP_total;
      vw.style.display = bad ? 'flex' : 'none';
      vw.innerHTML = bad ? '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M10.3 3.9L1.8 18a1.8 1.8 0 001.5 2.8h17.4a1.8 1.8 0 001.5-2.8L13.7 3.9a1.8 1.8 0 00-3.4 0z"/><path d="M12 9v4M12 17h.01"/></svg><span>One or more sources has fewer visits than bookings, which implies a conversion rate above 100% — double-check those numbers.</span>' : '';
    }
    const vt = $('#visits-totals');
    if (vt) {
      vt.innerHTML = '';
      vt.appendChild(el('span', {}, ['PS: ', el('b', { class: 'text-slate-800' }, [C.fmtNum(model.visitsPS_total)]), ' (' + C.fmtPct(model.convPS) + ')']));
      vt.appendChild(el('span', {}, ['Direct: ', el('b', { class: 'text-slate-800' }, [C.fmtNum(model.visitsDirect_total)]), ' (' + C.fmtPct(model.convDirect) + ')']));
      vt.appendChild(el('span', {}, ['CP: ', el('b', { class: 'text-slate-800' }, [C.fmtNum(model.visitsCP_total)]), ' (' + C.fmtPct(model.convCP) + ')']));
      vt.appendChild(el('span', {}, ['Total: ', el('b', { class: 'text-amber-600' }, [C.fmtNum(model.totalVisitsSum)]), ' (' + C.fmtPct(model.convOverall) + ')']));
    }
    const ht = $('#hoarding-total');
    if (ht) ht.textContent = 'Total Hoardings Cost: ' + C.fmtINR(model.hoardingsCost);
  }

  function renderHistoricalEditors() {
    const wrap = el('div');
    wrap.appendChild(el('p', { class: 'text-xs text-slate-500 mb-3' }, ['These feed the reporting slides (MTD tracking, lifetime history, agreement tracker, till-date advertising cost). Defaults are illustrative — edit any cell to reflect actuals.']));

    // MTD metrics
    wrap.appendChild(el('div', { class: 'text-[11px] font-semibold text-slate-500 mb-2 uppercase tracking-wide' }, ['MTD Report — ' + (state.months[0] || 'Month 1') + ', Budgeted vs Achieved']));
    const mtdTable = el('div', { class: 'overflow-x-auto mb-4' });
    const table = el('table', { class: 'mini-edit-table' });
    const thead = el('tr', {}, [el('th', {}, ['Metric']), ...state.mtdWeeks.flatMap(w => [el('th', {}, [w + ' Budget']), el('th', {}, [w + ' Achieved'])])]);
    table.appendChild(el('thead', {}, [thead]));
    const tbody = el('tbody');
    state.mtdMetrics.forEach((metric, mi) => {
      const tr = el('tr', {}, [el('td', { class: 'text-left text-xs text-slate-600' }, [metric.label])]);
      metric.budgeted.forEach((val, wi) => {
        const inp = numberInput(val, v => { state.mtdMetrics[mi].budgeted[wi] = v; recalc(); }, { step: 1 });
        tr.appendChild(el('td', {}, [inp]));
        const inp2 = numberInput(metric.achieved[wi], v => { state.mtdMetrics[mi].achieved[wi] = v; recalc(); }, { step: 1 });
        tr.appendChild(el('td', {}, [inp2]));
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    mtdTable.appendChild(table);
    wrap.appendChild(mtdTable);

    // Lifetime history
    wrap.appendChild(el('div', { class: 'text-[11px] font-semibold text-slate-500 mb-2 uppercase tracking-wide' }, ['Lifetime Visits & Booking Summary (Month-wise)']));
    const lifeTable = el('div', { class: 'overflow-x-auto mb-4' });
    const lt = el('table', { class: 'mini-edit-table' });
    lt.appendChild(el('thead', {}, [el('tr', {}, [el('th', {}, ['Month']), el('th', {}, ['Visits']), el('th', {}, ['Total Bookings']), el('th', {}, ['Live Bookings']), el('th', {}, ['Agreements'])])]));
    const ltbody = el('tbody');
    state.lifetimeHistory.forEach((row, ri) => {
      const tr = el('tr', {}, [
        el('td', {}, [(() => { const i = el('input', { class: 'text-input-sm', value: row.month }); i.addEventListener('input', e => { state.lifetimeHistory[ri].month = e.target.value; recalc(); }); return i; })()]),
        el('td', {}, [numberInput(row.visits, v => { state.lifetimeHistory[ri].visits = v; recalc(); })]),
        el('td', {}, [numberInput(row.totalBookings, v => { state.lifetimeHistory[ri].totalBookings = v; recalc(); })]),
        el('td', {}, [numberInput(row.liveBookings, v => { state.lifetimeHistory[ri].liveBookings = v; recalc(); })]),
        el('td', {}, [numberInput(row.agreements, v => { state.lifetimeHistory[ri].agreements = v; recalc(); })]),
      ]);
      ltbody.appendChild(tr);
    });
    lt.appendChild(ltbody);
    lifeTable.appendChild(lt);
    wrap.appendChild(lifeTable);

    // Agreement tracker
    wrap.appendChild(el('div', { class: 'text-[11px] font-semibold text-slate-500 mb-2 uppercase tracking-wide' }, ['Agreement Tracker']));
    const agTable = el('div', { class: 'overflow-x-auto mb-4' });
    const at = el('table', { class: 'mini-edit-table' });
    at.appendChild(el('thead', {}, [el('tr', {}, [el('th', {}, ['Sr']), el('th', {}, ['Client Name']), el('th', {}, ['Unit Number']), el('th', {}, ['Status & Remarks'])])]));
    const atbody = el('tbody');
    state.agreementTracker.forEach((row, ri) => {
      const tr = el('tr', {}, [
        el('td', { class: 'text-xs text-slate-500' }, [String(row.sr)]),
        el('td', {}, [(() => { const i = el('input', { class: 'text-input-sm', value: row.client }); i.addEventListener('input', e => { state.agreementTracker[ri].client = e.target.value; recalc(); }); return i; })()]),
        el('td', {}, [(() => { const i = el('input', { class: 'text-input-sm', value: row.unit }); i.addEventListener('input', e => { state.agreementTracker[ri].unit = e.target.value; recalc(); }); return i; })()]),
        el('td', {}, [(() => { const i = el('input', { class: 'text-input-sm w-full', value: row.status }); i.addEventListener('input', e => { state.agreementTracker[ri].status = e.target.value; recalc(); }); return i; })()]),
      ]);
      atbody.appendChild(tr);
    });
    at.appendChild(atbody);
    agTable.appendChild(at);
    wrap.appendChild(agTable);

    // Till-date advertising cost
    wrap.appendChild(el('div', { class: 'text-[11px] font-semibold text-slate-500 mb-2 uppercase tracking-wide' }, ['Till-Date Advertising Cost (Lifetime / Pre-April / Post-April)']));
    const tdGrid = el('div', { class: 'grid grid-cols-3 gap-4' });
    ['lifetime', 'preApril', 'postApril'].forEach(bucketKey => {
      const bucket = state.tillDateAdvertising[bucketKey];
      const label = bucketKey === 'lifetime' ? 'Lifetime' : bucketKey === 'preApril' ? 'Pre-April' : 'Post-April';
      const box = el('div', { class: 'p-3 rounded-lg border border-slate-200' }, [el('div', { class: 'text-xs font-bold text-slate-700 mb-2' }, [label])]);
      ['revenue', 'digital', 'ooh', 'branding', 'shoots', 'cpOutgo'].forEach(fieldKey => {
        const flabel = { revenue: 'Revenue', digital: 'Digital', ooh: 'OOH', branding: 'Branding', shoots: 'Shoots', cpOutgo: 'CP Outgo' }[fieldKey];
        box.appendChild(field(flabel, numberInput(bucket[fieldKey], v => { state.tillDateAdvertising[bucketKey][fieldKey] = v; recalc(); }, { step: 10000 })));
      });
      tdGrid.appendChild(box);
    });
    wrap.appendChild(tdGrid);

    return wrap;
  }

  /* ---------------------------- Deck Tab (17 slides) ---------------------------- */

  const SLIDE_NAMES = [
    'Cover', 'MTD Report', 'Lifetime Summary', 'Agreement Tracker', 'Till-Date Cost',
    'Way Forward Divider', 'Strategic Targets', 'Monthly Funnel (90d)', 'Digital Funnel',
    'Digital Bifurcation', 'Audience Matrix', 'OOH & Offline', 'CP Activation',
    'Creative Showcase', 'Festive Offers', '90-Day Budget', 'Closing',
  ];

  function slideShell(bodyHtml, opts) {
    opts = opts || {};
    const cls = 'slide-shell' + (opts.light ? ' light' : '') + (opts.blueprint ? ' blueprint-bg' : '');
    return `<div class="${cls}"><div class="slide-pad">${bodyHtml}</div></div>`;
  }

  function slideTitle(eyebrow, title) {
    return `<div class="mb-4"><div class="eyebrow">${eyebrow}</div><h2 class="slide-heading">${title}</h2><div class="ruler w-16 mt-2"></div></div>`;
  }

  function miniTable(head, rows, dark) {
    let html = `<table class="mini-table"><thead><tr>`;
    head.forEach((h, i) => { html += `<th style="background:${dark ? '#d97706' : '#0f172a'};text-align:${i === 0 ? 'left' : 'right'}">${h}</th>`; });
    html += `</tr></thead><tbody>`;
    rows.forEach((r, ri) => {
      html += `<tr style="background:${ri % 2 ? 'rgba(148,163,184,0.08)' : 'transparent'}">`;
      r.forEach((c, ci) => { html += `<td style="text-align:${ci === 0 ? 'left' : 'right'}">${c}</td>`; });
      html += `</tr>`;
    });
    html += `</tbody></table>`;
    return html;
  }

  function statChip(label, value) {
    return `<div class="stat-chip"><div class="stat-chip-label">${label}</div><div class="stat-chip-value">${value}</div></div>`;
  }

  function slideContent(idx) {
    const m = model, s = state;
    switch (idx) {
      case 0: // Cover
        return slideShell(`
          <div style="height:100%;display:flex;flex-direction:column;justify-content:center;align-items:flex-start;">
            <div class="eyebrow mb-4">${s.projectName} Way Forward Plan</div>
            <h1 class="cover-title">${s.months[0]} – ${s.months[m.n - 1]} ${s.year}</h1>
            <div class="ruler w-32 my-5"></div>
            <p style="color:#cbd5e1;font-size:1.1vw;max-width:55%;">A ${m.n}-month, data-engineered media plan built to convert ${C.fmtNum(m.totalVisitsSum)} site visits into ${m.totalBookings} confirmed bookings.</p>
            <div style="display:flex;gap:12px;margin-top:32px;">
              <div class="pill">₹${(m.totalRevenue / 1e7).toFixed(2)} Cr Revenue Target</div>
              <div class="pill">${m.totalBookings} Bookings</div>
              <div class="pill">${m.n} Months</div>
            </div>
          </div>`, { blueprint: true });

      case 1: { // MTD Report
        const rows = m.mtdMetrics.map(mm => [
          mm.label,
          ...mm.budgeted.flatMap((b, wi) => [C.fmtNum(b), C.fmtNum(mm.achieved[wi])]),
        ]);
        const head = ['Metric', ...s.mtdWeeks.flatMap(w => [w + ' Bud.', w + ' Ach.'])];
        return slideShell(slideTitle('Section 01', `${s.months[0]} Report — MTD Tracking`) + `<div style="overflow-x:auto;">${miniTable(head, rows)}</div>`);
      }

      case 2: { // Lifetime visits & booking summary
        const rows = m.lifetimeHistory.map(r => [r.month, C.fmtNum(r.visits), C.fmtNum(r.totalBookings), C.fmtNum(r.liveBookings), C.fmtNum(r.agreements)]);
        return slideShell(slideTitle('Section 02', 'Lifetime Visits & Booking Summary — Month-wise') +
          miniTable(['Month', 'Visits', 'Total Bookings', 'Live Bookings', 'Agreements'], rows) +
          `<div style="display:flex;gap:12px;margin-top:16px;">${statChip('Total Visits', C.fmtNum(m.lifetimeTotals.visits))}${statChip('Total Bookings', C.fmtNum(m.lifetimeTotals.totalBookings))}${statChip('Live Bookings', C.fmtNum(m.lifetimeTotals.liveBookings))}${statChip('Agreements Done', C.fmtNum(m.lifetimeTotals.agreements))}</div>`,
          { light: true });
      }

      case 3: { // Agreement tracker
        const rows = m.agreementTracker.map(r => [String(r.sr), r.client, r.unit, r.status]);
        return slideShell(slideTitle('Section 03', 'Agreement Tracker with Remarks') + miniTable(['Sr No.', 'Client Name', 'Unit Number', 'Agreement Status & Remarks'], rows));
      }

      case 4: { // Till date advertising cost
        const td = m.tillDate;
        const cols = ['lifetime', 'preApril', 'postApril'];
        const labels = { lifetime: 'Lifetime', preApril: 'Pre-April', postApril: 'Post-April' };
        const rowsDef = [
          ['Revenue', b => C.fmtINR(b.revenue, true)],
          ['Digital', b => C.fmtINR(b.digital, true)],
          ['OOH', b => C.fmtINR(b.ooh, true)],
          ['Branding', b => C.fmtINR(b.branding, true)],
          ['Shoots', b => C.fmtINR(b.shoots, true)],
          ['CP Outgo', b => C.fmtINR(b.cpOutgo, true)],
          ['Total Marketing Spend', b => C.fmtINR(b.totalSpend, true)],
          ['Advertising Cost %', b => C.fmtPct(b.costPct)],
        ];
        const rows = rowsDef.map(([label, fn]) => [label, ...cols.map(c => fn(td[c]))]);
        return slideShell(slideTitle('Section 04', 'Till-Date Advertising Cost Calculation') + miniTable(['Metric', labels.lifetime, labels.preApril, labels.postApril], rows), { light: true });
      }

      case 5: // Divider
        return slideShell(`
          <div style="height:100%;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;">
            <div class="eyebrow mb-3">Section 05</div>
            <h1 style="font-size:3.2vw;font-weight:800;font-family:Georgia,serif;">Way Forward Strategy</h1>
            <div class="ruler w-24 my-5"></div>
            <div style="font-size:2vw;color:#f59e0b;">→</div>
          </div>`, { blueprint: true });

      case 6: { // Strategic targets & channel attribution
        const channelRows = [
          ['Digital / Presales', C.fmtNum(m.totalLeadsSum), C.fmtNum(m.visitsPS_total), String(m.bookingsPS_total)],
          ['Direct / Walk-in', '—', C.fmtNum(m.visitsDirect_total), String(m.bookingsDirect_total)],
          ['Channel Partner', '—', C.fmtNum(m.visitsCP_total), String(m.bookingsCP_total)],
        ];
        return slideShell(slideTitle('Section 06', 'Strategic Targets & Channel Attribution') + `
          <div style="display:flex;gap:12px;margin-bottom:20px;">
            ${statChip('Duration', m.n + ' Months')}
            ${statChip('Total Bookings', String(m.totalBookings))}
            ${statChip('Digital + Direct', String(m.bookingsPS_total + m.bookingsDirect_total))}
            ${statChip('Channel Partner', String(m.bookingsCP_total))}
          </div>
          ${miniTable(['Tool', 'Leads', 'Estimated Walk-ins', 'Bookings'], channelRows, true)}
        `);
      }

      case 7: { // Monthly funnel targets (90 days)
        const rows = m.months.map((mo, i) => [
          mo, C.fmtINR(m.totalDigitalSpend[i], true), C.fmtNum(m.totalDigitalLeads[i]), C.fmtINR(m.cpl[i]),
          C.fmtNum(m.qualifiedLeads[i]), C.fmtINR(m.cpqlArr[i]), C.fmtNum(m.presalesSVS[i]), C.fmtNum(m.visitsPS[i]),
          C.fmtNum(m.visitsDirect[i]), C.fmtNum(m.visitsCP[i]), C.fmtNum(m.revisits[i]), String(m.monthBookings[i]),
        ]);
        return slideShell(slideTitle('Section 07', 'Month-Wise Funnel Targets (90 Days)') +
          `<div style="overflow-x:auto;">${miniTable(['Month', 'Spend', 'Leads', 'CPL', 'Qual. Leads', 'CPQL', 'SVS', 'PS Visits', 'Walk-ins', 'CP Visits', 'Revisits', 'Bookings'], rows)}</div>`, { light: true });
      }

      case 8: { // Digital funnel breakdown & objective/platform split
        const funnel = [['Total Digital Leads', m.totalLeadsSum], ['Qualified Leads', m.totalQualifiedLeads], ['Site Visits (Presales)', m.visitsPS_total], ['Bookings (Presales)', m.bookingsPS_total]];
        const max = funnel[0][1] || 1;
        const bars = funnel.map(([label, val]) => `
          <div style="margin-bottom:10px;">
            <div style="display:flex;justify-content:space-between;font-size:0.78vw;margin-bottom:4px;"><span>${label}</span><b>${C.fmtNum(val)}</b></div>
            <div style="height:10px;border-radius:6px;background:#f1f5f9;overflow:hidden;"><div style="height:100%;border-radius:6px;width:${Math.max(4, (val / max) * 100)}%;background:linear-gradient(90deg,#d97706,#f59e0b);"></div></div>
          </div>`).join('');
        const objRows = Object.values(m.objectives).map(o => [o.label, C.fmtPct(o.leadPct), C.fmtNum(o.leads), C.fmtPct(o.budgetPct), C.fmtINR(o.budget, true)]);
        const platRows = m.platforms.map(p => [p.label, p.pct ? C.fmtPct(p.pct) : '—', p.cpl ? C.fmtINR(p.cpl) : '—', C.fmtNum(p.leads), p.budget ? C.fmtINR(p.budget, true) : '₹0']);
        return slideShell(slideTitle('Section 08', 'Digital Funnel Breakdown & Allocation') + `
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:24px;">
            <div>
              <div class="mini-heading">Conversion Cascade</div>
              ${bars}
            </div>
            <div>
              <div class="mini-heading">Digital Objective Split</div>
              ${miniTable(['Objective', 'Leads %', 'Leads', 'Budget %', 'Budget'], objRows, true)}
              <div class="mini-heading" style="margin-top:10px;">Platform Budget Bifurcation</div>
              ${miniTable(['Platform', 'Wt.', 'CPL', 'Leads', 'Budget'], platRows, true)}
            </div>
          </div>`, { light: true });
      }

      case 9: { // Digital detailed bifurcation (campaign matrix)
        const rows = m.campaigns.map(c => [c.name, c.objective, c.audience, C.fmtINR(c.budget, true), C.fmtPct(c.pct)]);
        return slideShell(slideTitle('Section 09', 'Digital Detailed Bifurcation — Campaign Matrix') + `<div style="overflow-x:auto;">${miniTable(['Campaign', 'Objective', 'Audience Targeting', 'Monthly Budget', '% Budget'], rows)}</div>`);
      }

      case 10: { // Audience & messaging framework
        const tiers = [
          { tag: '1st Touch — Lead Generation', demo: 'HNI & upgrade buyers, 32–55', geo: 'City core, premium suburbs, NRI corridors', medium: 'Meta, Google Search, Pmax', angle: 'Aspiration + scarcity — limited inventory, skyline living', content: 'Cinematic walkthroughs, amenity films, price reveal' },
          { tag: '2nd Touch — Remarketing', demo: 'Engaged leads & website visitors', geo: 'Retarget pools from 1st-touch campaigns', medium: 'Meta remarketing, Google Display, WhatsApp', angle: 'Urgency + social proof — offers, testimonials', content: 'Testimonial reels, offer carousels, CP referral pitch' },
        ];
        const cards = tiers.map(t => `
          <div class="persona-card">
            <div class="persona-tag">${t.tag}</div>
            <div class="persona-row"><b>Demographics</b><span>${t.demo}</span></div>
            <div class="persona-row"><b>Target Geos</b><span>${t.geo}</span></div>
            <div class="persona-row"><b>Mediums</b><span>${t.medium}</span></div>
            <div class="persona-row"><b>Communication Angle</b><span>${t.angle}</span></div>
            <div class="persona-row"><b>Content Buckets</b><span>${t.content}</span></div>
          </div>`).join('');
        return slideShell(slideTitle('Section 10', 'Audience Targeting & Messaging Framework') + `<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">${cards}</div>`, { light: true });
      }

      case 11: { // OOH & offline
        const rows = m.oohItems.map(o => [o.label, C.fmtINR(o.value)]);
        return slideShell(slideTitle('Section 11', 'OOH & Offline Activation Split') + miniTable(['Marketing Tools', 'Cost'], rows) +
          `<div style="display:flex;gap:12px;margin-top:16px;">${statChip('Total Offline Investment', C.fmtINR(m.oohItems.reduce((a, o) => a + o.value, 0), true))}</div>`);
      }

      case 12: { // CP activation targets
        const rows = m.months.map((mo, i) => [mo, String(m.cpDigitalReach[i]), C.fmtNum(m.outreach[i]), C.fmtNum(m.physicalMeetings[i]), String(m.catA[i]), String(m.catB[i]), C.fmtNum(m.visitsCP[i]), String(m.bCP[i])]);
        return slideShell(slideTitle('Section 12', 'Channel Partner Activation Targets') + `<div style="overflow-x:auto;">${miniTable(['Month', 'Digital Reach', 'WA/Email Outreach', 'Physical Meetings', 'CAT A', 'CAT B', 'CP Visits', 'Conversions'], rows, true)}</div>`, { light: true });
      }

      case 13: // Creative showcase
        return slideShell(slideTitle('Section 13', 'Creative Showcase — Campaign Formats') + `
          <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:16px;">
            ${['Location Advantage', 'Lifestyle Amenities', 'Skyline & Interiors', 'Developer Trust Films'].map(t => `
              <div class="creative-card">
                <div class="creative-thumb"></div>
                <div class="creative-label">${t}</div>
              </div>`).join('')}
          </div>`);

      case 14: // Festive offers
        return slideShell(slideTitle('Section 14', 'Festive Offers & CP Referral Schemes') + `
          <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:16px;">
            ${['Independence Day Offer', 'Festive Season Bonanza', 'Early-Bird Booking Benefits', 'CP Referral Scheme'].map(t => `
              <div class="creative-card gold">
                <div class="creative-thumb gold"></div>
                <div class="creative-label">${t}</div>
              </div>`).join('')}
          </div>`, { light: true });

      case 15: { // 90-day budget bifurcation
        const rows = m.expenses.map(e => [e.label, C.fmtINR(e.value)]);
        return slideShell(slideTitle('Section 15', 'Comprehensive 90-Day Media Plan & Cost Bifurcation') + `
          <div style="display:grid;grid-template-columns:2fr 1fr;gap:24px;">
            <div>${miniTable(['Line Item', 'Investment'], rows)}</div>
            <div style="display:flex;flex-direction:column;gap:10px;">
              ${statChip('Total Advertising Budget', C.fmtINR(m.totalAdvertisingBudget, true))}
              ${statChip('Total Revenue', C.fmtINR(m.totalRevenue, true))}
              ${statChip('Advertising Cost %', C.fmtPct(m.advertisingCostPct))}
              ${statChip('Net Marketing Cost %', C.fmtPct(m.totalMarketingCostPct))}
            </div>
          </div>`);
      }

      case 16: // Closing
        return slideShell(`
          <div style="height:100%;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;">
            <div class="eyebrow mb-3">Thank You</div>
            <h1 style="font-size:3vw;font-weight:800;font-family:Georgia,serif;">${s.projectName}</h1>
            <div class="ruler w-24 my-5"></div>
            <p style="color:#cbd5e1;font-size:1vw;">For queries, reach the marketing team.</p>
            <p style="color:#94a3b8;font-size:0.8vw;margin-top:6px;">[ Contact placeholder ] · [ Email placeholder ] · [ Phone placeholder ]</p>
          </div>`, { blueprint: true });

      default:
        return slideShell('<div>Slide not found</div>');
    }
  }

  function renderDeck() {
    const pillsWrap = $('#slide-pills');
    if (pillsWrap && pillsWrap.children.length !== SLIDE_NAMES.length) {
      pillsWrap.innerHTML = '';
      SLIDE_NAMES.forEach((name, i) => {
        const btn = el('button', { type: 'button', class: 'slide-pill' + (i === currentSlide ? ' active' : '') }, [`${i + 1}. ${name}`]);
        btn.addEventListener('click', () => { currentSlide = i; renderDeck(); });
        pillsWrap.appendChild(btn);
      });
    } else if (pillsWrap) {
      $$('.slide-pill').forEach((btn, i) => btn.classList.toggle('active', i === currentSlide));
    }

    const stage = $('#slide-stage');
    if (stage) stage.innerHTML = slideContent(currentSlide);

    const dots = $('#slide-dots');
    if (dots) {
      dots.innerHTML = '';
      SLIDE_NAMES.forEach((_, i) => {
        const d = el('button', { type: 'button', class: 'slide-dot' + (i === currentSlide ? ' active' : '') });
        d.addEventListener('click', () => { currentSlide = i; renderDeck(); });
        dots.appendChild(d);
      });
    }
  }

  function goSlide(delta) {
    currentSlide = Math.min(SLIDE_NAMES.length - 1, Math.max(0, currentSlide + delta));
    renderDeck();
  }

  /* ---------------------------- Export Tab ---------------------------- */

  function renderExportSummary() {
    const root = $('#export-summary');
    if (!root) return;
    root.innerHTML = `
      <div class="grid grid-cols-4 gap-3">
        ${statBadge('Total Bookings', C.fmtNum(model.totalBookings))}
        ${statBadge('Total Revenue', C.fmtINR(model.totalRevenue, true))}
        ${statBadge('Ad Budget', C.fmtINR(model.totalAdvertisingBudget, true))}
        ${statBadge('Marketing Cost %', C.fmtPct(model.totalMarketingCostPct))}
      </div>`;
    renderCharts();
  }

  function statBadge(label, value) {
    return `<div class="stat-badge"><div class="stat-badge-label">${label}</div><div class="stat-badge-value">${value}</div></div>`;
  }

  let funnelChart = null, budgetChart = null;
  function renderCharts() {
    const funnelCanvas = $('#funnel-chart');
    const budgetCanvas = $('#budget-chart');
    if (!funnelCanvas || !budgetCanvas || typeof Chart === 'undefined') return;
    try {
      if (funnelChart) funnelChart.destroy();
      funnelChart = new Chart(funnelCanvas.getContext('2d'), {
        type: 'bar',
        data: {
          labels: ['Digital Leads', 'Qualified Leads', 'Site Visits', 'Bookings'],
          datasets: [{ data: [model.totalLeadsSum, model.totalQualifiedLeads, model.totalVisitsSum, model.totalBookings], backgroundColor: ['#0f172a', '#334155', '#d97706', '#f59e0b'], borderRadius: 6 }],
        },
        options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { grid: { display: false } } } },
      });
    } catch (e) { console.error('Funnel chart failed:', e); }
    try {
      if (budgetChart) budgetChart.destroy();
      budgetChart = new Chart(budgetCanvas.getContext('2d'), {
        type: 'doughnut',
        data: { labels: model.expenses.map(e => e.label), datasets: [{ data: model.expenses.map(e => e.value), backgroundColor: ['#0f172a', '#1e293b', '#334155', '#475569', '#64748b', '#94a3b8', '#d97706', '#f59e0b', '#fbbf24', '#fcd34d', '#78716c', '#a8a29e', '#e7e5e4'] }] },
        options: { plugins: { legend: { display: false } } },
      });
    } catch (e) { console.error('Budget chart failed:', e); }
  }

  /* ---------------------------- Tabs & Sidebar ---------------------------- */

  function updateSidebarSummary() {
    const b = $('#sidebar-bookings'), r = $('#sidebar-revenue'), d = $('#sidebar-duration'), p = $('#sidebar-project-name');
    if (b) b.textContent = String(model.totalBookings);
    if (r) r.textContent = C.fmtINR(model.totalRevenue, true);
    if (d) d.textContent = model.n + ' mo';
    if (p) p.textContent = state.projectName;
    if ($('#visits-warning')) updateComputedBits();
  }

  function switchTab(tab) {
    try {
      currentTab = tab;
      $$('.tab-panel').forEach(p => p.classList.toggle('active', p.id === 'tab-' + tab));
      $$('.nav-tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
      if (tab === 'deck') renderDeck();
      if (tab === 'export') renderExportSummary();
    } catch (e) {
      console.error('Tab switch failed:', e);
      showFatalError(e);
    }
  }

  function showFatalError(e) {
    const box = document.getElementById('boot-error');
    if (!box) return;
    box.style.display = 'block';
    box.innerHTML = '<h2 style="color:#f59e0b;margin-bottom:12px;">Something went wrong while rendering</h2>' +
      '<p style="margin-bottom:8px;font-size:13px;color:#cbd5e1;">' + (e && (e.message || String(e))) + '</p>' +
      '<button onclick="location.reload()" style="margin-top:16px;padding:10px 20px;background:#d97706;color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:600;">Reload</button>';
  }

  /* ---------------------------- Icons ---------------------------- */

  const ICONS = {
    building: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 21V6l8-3 8 3v15"/><path d="M4 21h16"/><path d="M10 21v-4h4v4"/></svg>',
    target: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1"/></svg>',
    mapPin: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 21s7-6.5 7-12a7 7 0 10-14 0c0 5.5 7 12 7 12z"/><circle cx="12" cy="9" r="2.5"/></svg>',
    sliders: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M2 14h4M10 8h4M18 12h4"/></svg>',
    megaphone: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M3 11v2a2 2 0 002 2h1l3 5V6l-3 5H5a2 2 0 00-2 2z"/><path d="M14 8a4 4 0 010 8"/><path d="M17 5a8 8 0 010 14"/></svg>',
    users: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="9" cy="8" r="3"/><path d="M2 21v-1a5 5 0 015-5h4a5 5 0 015 5v1"/><circle cx="17" cy="8" r="2.5"/></svg>',
    calendar: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>',
  };

  /* ---------------------------- Init ---------------------------- */

  function init() {
    $$('.nav-tab-btn').forEach(btn => btn.addEventListener('click', () => switchTab(btn.dataset.tab)));
    $('#slide-prev').addEventListener('click', () => goSlide(-1));
    $('#slide-next').addEventListener('click', () => goSlide(1));
    $('#slide-fullscreen').addEventListener('click', () => {
      const wrap = $('#deck-stage-wrap');
      if (!document.fullscreenElement) { wrap.requestFullscreen && wrap.requestFullscreen(); }
      else { document.exitFullscreen && document.exitFullscreen(); }
    });
    window.addEventListener('keydown', e => {
      if (currentTab !== 'deck') return;
      if (e.key === 'ArrowRight') goSlide(1);
      if (e.key === 'ArrowLeft') goSlide(-1);
    });

    $('#export-btn').addEventListener('click', async () => {
      const btn = $('#export-btn');
      const original = btn.textContent;
      btn.textContent = 'Building Deck…';
      btn.disabled = true;
      try {
        await window.PptxGenerator.generate(state, model);
      } catch (e) {
        console.error(e);
        alert('Export failed: ' + e.message);
      }
      btn.textContent = original;
      btn.disabled = false;
    });

    renderPlanner();
    switchTab('planner');
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => { try { init(); } catch (e) { console.error(e); showFatalError(e); } });
  } else {
    try { init(); } catch (e) { console.error(e); showFatalError(e); }
  }

  // Expose for debugging / tests
  window.__app = { getState: () => state, getModel: () => model, setState, recalc };

})();
