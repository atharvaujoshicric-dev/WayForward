/* ==========================================================================
   calculator.js
   Pure calculation module for the Real Estate Media Plan Calculator.
   No DOM access here — every function takes plain data in and returns
   plain data out, so it can be unit-tested or reused by the PPTX generator
   and the slide-deck renderer without duplication.
   ========================================================================== */

(function (global) {
  'use strict';

  /* ---------------------------- Utilities ---------------------------- */

  function sum(arr) {
    return (arr || []).reduce((a, b) => a + (Number(b) || 0), 0);
  }

  function clampArr(arr, n, fill) {
    fill = fill === undefined ? 0 : fill;
    const out = (arr || []).slice(0, n);
    while (out.length < n) out.push(fill);
    return out;
  }

  function round(n) {
    return Math.round(Number(n) || 0);
  }

  function fmtINR(n, compact) {
    n = Number(n) || 0;
    if (compact) {
      const abs = Math.abs(n);
      if (abs >= 1e7) return '₹' + (n / 1e7).toFixed(2) + ' Cr';
      if (abs >= 1e5) return '₹' + (n / 1e5).toFixed(2) + ' L';
      if (abs >= 1e3) return '₹' + (n / 1e3).toFixed(1) + ' K';
      return '₹' + Math.round(n).toLocaleString('en-IN');
    }
    return '₹' + Math.round(n).toLocaleString('en-IN');
  }

  function fmtNum(n, d) {
    d = d || 0;
    n = Number(n) || 0;
    return n.toLocaleString('en-IN', { maximumFractionDigits: d, minimumFractionDigits: d });
  }

  function fmtPct(n, d) {
    d = d === undefined ? 1 : d;
    n = Number(n) || 0;
    return (n * 100).toFixed(d) + '%';
  }

  function deriveVisits(bookingsArr, conv) {
    const c = conv > 0 ? conv : 0.09;
    return (bookingsArr || []).map(b => round((Number(b) || 0) / c));
  }

  /* ---------------------------- Default / Preset State ---------------------------- */

  function makeState(overrides) {
    overrides = overrides || {};
    const base = {
      projectName: 'SKYLUXE',
      months: ['August', 'September', 'October'],
      year: 2026,
      duration: 3,
      avgTicketSize: 9700000,
      conversionRate: 0.09,

      // Forward 90-day plan — bookings target & visit split (both fully editable)
      bookingsPS: [3, 4, 5],
      bookingsDirect: [3, 2, 3],
      bookingsCP: [7, 8, 8],

      visitsPS: [33, 44, 56],
      visitsDirect: [33, 22, 33],
      visitsCP: [78, 89, 89],

      cplMonthly: [1200, 1150, 1150],

      platformPct: { meta: 0.625, pmax: 0.15, search: 0.10, experimental: 0.075, nri: 0.05 },
      platformCPL: { meta: 1096, pmax: 1632, search: 1265, experimental: 1855, nri: 2134 },

      // Campaign-level bifurcation (Slide 10) — % of Digital Lead-Gen budget
      campaigns: [
        { name: 'Meta — Lookalike Audiences', objective: 'Lead Generation', audience: 'LAL 1-3% of CRM & website converters', pct: 0.22 },
        { name: 'Meta — Interest Based', objective: 'Lead Generation', audience: 'HNI interests: luxury, real estate, travel', pct: 0.19 },
        { name: 'Meta — Hyperlocal Pin Drop', objective: 'Lead Generation', audience: '5–10km radius of project site', pct: 0.14 },
        { name: 'Meta — Generic Broad', objective: 'Awareness + Leads', audience: 'Broad 28–55, city-wide', pct: 0.13 },
        { name: 'Google Search — Primary Keywords', objective: 'Lead Generation', audience: 'High-intent project & micro-market search', pct: 0.12 },
        { name: 'Google Search — Secondary Keywords', objective: 'Lead Generation', audience: 'Category & competitor search terms', pct: 0.06 },
        { name: 'Google Discovery / Pmax', objective: 'Lead Generation', audience: 'In-market real estate audiences', pct: 0.06 },
        { name: 'Remarketing — All Platforms', objective: 'Remarketing', audience: 'Website visitors & engaged leads', pct: 0.03 },
        { name: 'WhatsApp Click-to-Chat', objective: 'Lead Generation', audience: 'Mobile-first high-intent audience', pct: 0.03 },
      ],

      socialMediaBoosting: 45000,
      droneShoot: 75000,
      ugcVideos: 50000,
      developerByteShoot: 45000,
      chatbotWhatsapp: 70000,
      societyMallActivations: 250000,
      hoardings: { count: 4, months: 3, rate: 95000 },
      kiosks: 90000,
      newspaperLeaflet: 150000,
      cpGathering: 110000,
      cpHighTea: 60000,
      cpBrokeragePercent: 0.03,
      cpVisitIncentive: 600,
      cpBookingIncentive: 25000,

      catA: [12, 15, 15],
      catB: [22, 24, 26],
      cpDigitalReach: [650, 720, 750],

      /* ---------------- Historical / Reporting data (Slides 2–5) ---------------- */

      // Slide 2 — MTD report for the first month, 4-week tracking, budgeted vs achieved
      mtdWeeks: ['Week 1', 'Week 2', 'Week 3', 'Week 4'],
      mtdMetrics: [
        { key: 'spend', label: 'Marketing Spends (₹)', budgeted: [220000, 230000, 210000, 215000], achieved: [205000, 218000, 200000, 209000] },
        { key: 'branding', label: 'Branding Spends (₹)', budgeted: [26000, 27000, 25000, 26000], achieved: [24000, 25500, 23500, 24500] },
        { key: 'digitalLG', label: 'Digital LG Spends (₹)', budgeted: [194000, 203000, 185000, 189000], achieved: [181000, 192500, 176500, 184500] },
        { key: 'leads', label: 'Leads', budgeted: [165, 172, 158, 161], achieved: [154, 163, 150, 156] },
        { key: 'cpl', label: 'CPL (₹)', budgeted: [1176, 1180, 1171, 1174], achieved: [1175, 1181, 1177, 1183] },
        { key: 'qualifiedLeads', label: 'Qualified Leads', budgeted: [63, 65, 60, 61], achieved: [58, 62, 56, 59] },
        { key: 'cpql', label: 'CPQL (₹)', budgeted: [3079, 3123, 3083, 3098], achieved: [3121, 3105, 3152, 3128] },
        { key: 'svs', label: 'Presales SVS', budgeted: [31, 33, 30, 30], achieved: [29, 31, 28, 29] },
        { key: 'visits', label: 'Total Visits', budgeted: [37, 38, 35, 36], achieved: [34, 36, 33, 35] },
        { key: 'revisits', label: 'Revisits', budgeted: [8, 8, 8, 8], achieved: [7, 8, 7, 8] },
        { key: 'bookings', label: 'Bookings', budgeted: [3, 3, 3, 4], achieved: [2, 3, 2, 3] },
      ],

      // Slide 3 — Lifetime visits & booking summary, month-wise, Dec-25 to current
      lifetimeHistory: [
        { month: 'Dec-25', visits: 210, totalBookings: 22, liveBookings: 18, agreements: 15 },
        { month: 'Jan-26', visits: 245, totalBookings: 26, liveBookings: 21, agreements: 19 },
        { month: 'Feb-26', visits: 268, totalBookings: 29, liveBookings: 24, agreements: 22 },
        { month: 'Mar-26', visits: 252, totalBookings: 27, liveBookings: 23, agreements: 20 },
        { month: 'Apr-26', visits: 231, totalBookings: 24, liveBookings: 20, agreements: 18 },
        { month: 'May-26', visits: 219, totalBookings: 23, liveBookings: 19, agreements: 17 },
        { month: 'Jun-26', visits: 238, totalBookings: 25, liveBookings: 21, agreements: 18 },
        { month: 'Jul-26', visits: 226, totalBookings: 24, liveBookings: 20, agreements: 18 },
      ],

      // Slide 4 — Agreement tracker with remarks
      agreementTracker: [
        { sr: 1, client: 'R. Mehta', unit: 'A-1204', status: 'Agreement registered — 12 Jul 2026' },
        { sr: 2, client: 'S. & P. Nair', unit: 'B-0803', status: 'Documentation in progress, expected 25 Aug' },
        { sr: 3, client: 'A. Kapoor', unit: 'A-1601', status: 'Token received, agreement drafting stage' },
        { sr: 4, client: 'V. Reddy', unit: 'C-0405', status: 'Home loan sanction awaited' },
        { sr: 5, client: 'M. Shah Family Trust', unit: 'B-1102', status: 'Agreement registered — 03 Aug 2026' },
      ],

      // Slide 5 — Till-date advertising cost calculation (Lifetime / Pre-April / Post-April)
      tillDateAdvertising: {
        lifetime: { revenue: 184300000, digital: 5200000, ooh: 3400000, branding: 1800000, shoots: 950000, cpOutgo: 2100000 },
        preApril: { revenue: 106700000, digital: 3100000, ooh: 2200000, branding: 1050000, shoots: 620000, cpOutgo: 1250000 },
        postApril: { revenue: 77600000, digital: 2100000, ooh: 1200000, branding: 750000, shoots: 330000, cpOutgo: 850000 },
      },
    };
    return Object.assign({}, base, overrides);
  }

  const SKYLUXE_PRESET = makeState();

  function fastTrackPreset() {
    const bookingsPS = [2, 2, 2];
    const bookingsDirect = [1, 1, 2];
    const bookingsCP = [3, 3, 4];
    const conv = 0.09;
    return makeState({
      projectName: 'SKYLUXE — FAST TRACK',
      duration: 3,
      months: ['August', 'September', 'October'],
      bookingsPS: bookingsPS,
      bookingsDirect: bookingsDirect,
      bookingsCP: bookingsCP,
      visitsPS: deriveVisits(bookingsPS, conv),
      visitsDirect: deriveVisits(bookingsDirect, conv),
      visitsCP: deriveVisits(bookingsCP, conv),
      cplMonthly: [1150, 1100, 1100],
      catA: [8, 10, 10],
      catB: [16, 18, 20],
      cpDigitalReach: [500, 560, 600],
    });
  }

  function resizeStateToDuration(state, n, months) {
    n = Math.max(1, Math.min(12, Number(n) || 1));
    const grow = (arr, fill) => {
      const out = (arr || []).slice(0, n);
      while (out.length < n) out.push(fill);
      return out;
    };
    const monthNames = (months && months.length)
      ? clampArr(months, n, 'Month ' + (months.length + 1))
      : grow(state.months, 'Month ' + (state.months.length + 1));
    return Object.assign({}, state, {
      duration: n,
      months: monthNames,
      bookingsPS: grow(state.bookingsPS, 0),
      bookingsDirect: grow(state.bookingsDirect, 0),
      bookingsCP: grow(state.bookingsCP, 0),
      visitsPS: grow(state.visitsPS, 0),
      visitsDirect: grow(state.visitsDirect, 0),
      visitsCP: grow(state.visitsCP, 0),
      cplMonthly: grow(state.cplMonthly, state.cplMonthly[state.cplMonthly.length - 1] || 1000),
      catA: grow(state.catA, 0),
      catB: grow(state.catB, 0),
      cpDigitalReach: grow(state.cpDigitalReach, 0),
    });
  }

  /* ---------------------------- Core Calculation Engine ---------------------------- */

  function computeModel(s) {
    const n = Math.max(1, Math.min(12, Number(s.duration) || 1));
    const months = clampArr(s.months, n, 'Month');

    const bPS = clampArr(s.bookingsPS, n);
    const bDirect = clampArr(s.bookingsDirect, n);
    const bCP = clampArr(s.bookingsCP, n);
    const cpl = clampArr(s.cplMonthly, n, s.cplMonthly[s.cplMonthly.length - 1] || 1000);

    // Site visits are directly user-editable (per source, per month).
    // Conversion % below is computed FROM these for reporting/insight only.
    const visitsPS = clampArr(s.visitsPS, n);
    const visitsDirect = clampArr(s.visitsDirect, n);
    const visitsCP = clampArr(s.visitsCP, n);
    const totalVisits = months.map((_, i) => visitsPS[i] + visitsDirect[i] + visitsCP[i]);
    const revisits = totalVisits.map(v => v * 0.22);

    const presalesSVS = visitsPS.map(v => v * 2);
    const qualifiedLeads = presalesSVS.map(v => v * 2);
    const totalDigitalLeads = qualifiedLeads.map(q => q / 0.38);

    const digitalLGSpend = totalDigitalLeads.map((l, i) => l * (cpl[i] || 0));
    const brandingSpend = digitalLGSpend.map(v => v * 0.12);
    const totalDigitalSpend = digitalLGSpend.map((v, i) => v + brandingSpend[i]);
    const cpqlArr = digitalLGSpend.map((v, i) => (qualifiedLeads[i] > 0 ? v / qualifiedLeads[i] : 0));
    const monthBookings = months.map((_, i) => bPS[i] + bDirect[i] + bCP[i]);

    const bookingsPS_total = sum(bPS);
    const bookingsDirect_total = sum(bDirect);
    const bookingsCP_total = sum(bCP);
    const totalBookings = bookingsPS_total + bookingsDirect_total + bookingsCP_total;

    const totalVisitsSum = sum(totalVisits);
    const totalRevisitsSum = sum(revisits);
    const totalQualifiedLeads = sum(qualifiedLeads);
    const totalLeadsSum = sum(totalDigitalLeads);
    const totalDigitalLGSpendSum = sum(digitalLGSpend);
    const totalBrandingSpendSum = sum(brandingSpend);
    const totalDigitalSpendSum = totalDigitalLGSpendSum + totalBrandingSpendSum;
    const totalSVS = sum(presalesSVS);

    const visitsPS_total = sum(visitsPS), visitsDirect_total = sum(visitsDirect), visitsCP_total = sum(visitsCP);
    const convPS = visitsPS_total > 0 ? bookingsPS_total / visitsPS_total : 0;
    const convDirect = visitsDirect_total > 0 ? bookingsDirect_total / visitsDirect_total : 0;
    const convCP = visitsCP_total > 0 ? bookingsCP_total / visitsCP_total : 0;
    const convOverall = totalVisitsSum > 0 ? totalBookings / totalVisitsSum : 0;

    // Digital objective breakdown
    const objectives = {
      leadGen: { label: 'Lead Generation', leadPct: 0.83, budgetPct: 0.85, leads: totalLeadsSum * 0.83, budget: totalDigitalLGSpendSum * 0.85 },
      branding: { label: 'Branding', leadPct: 0, budgetPct: 1.0, leads: 0, budget: totalBrandingSpendSum * 1.0 },
      remarketing: { label: 'Remarketing', leadPct: 0.105, budgetPct: 0.15, leads: totalLeadsSum * 0.105, budget: totalDigitalLGSpendSum * 0.15 },
      organic: { label: 'Organic', leadPct: 0.065, budgetPct: 0, leads: totalLeadsSum * 0.065, budget: 0 },
    };

    // Platform allocation
    const platformDefs = [
      { key: 'meta', label: 'Meta / Facebook', pct: s.platformPct.meta, cpl: s.platformCPL.meta, audience: 'Broad + lookalike, 25–55' },
      { key: 'pmax', label: 'Google Pmax + Demand Gen', pct: s.platformPct.pmax, cpl: s.platformCPL.pmax, audience: 'High-intent search & display' },
      { key: 'search', label: 'Google Search', pct: s.platformPct.search, cpl: s.platformCPL.search, audience: 'Bottom-funnel keyword intent' },
      { key: 'experimental', label: 'Experimental / LinkedIn / Native', pct: s.platformPct.experimental, cpl: s.platformCPL.experimental, audience: 'Affluent LinkedIn / native' },
      { key: 'nri', label: 'NRI Campaigns — USA/GCC', pct: s.platformPct.nri, cpl: s.platformCPL.nri, audience: 'NRI corridors — USA / GCC' },
    ];
    const platforms = platformDefs.map(p => {
      const budget = totalDigitalLGSpendSum * (Number(p.pct) || 0);
      const leads = (Number(p.cpl) || 0) > 0 ? budget / p.cpl : 0;
      return Object.assign({}, p, { budget: budget, leads: leads });
    });
    platforms.push({ key: 'organic', label: 'Organic (High-Intent Search)', pct: 0, cpl: 0, budget: 0, leads: objectives.organic.leads, audience: 'Organic + direct search' });

    // Campaign-level bifurcation (Slide 10)
    const campaigns = (s.campaigns || []).map(c => ({
      name: c.name,
      objective: c.objective,
      audience: c.audience,
      pct: c.pct,
      budget: totalDigitalLGSpendSum * (Number(c.pct) || 0),
    }));

    // Offline / OOH expenses
    const hoardingsCost = (Number(s.hoardings.count) || 0) * (Number(s.hoardings.months) || 0) * (Number(s.hoardings.rate) || 0);
    const cpVisitsTotal = sum(visitsCP);
    const cpIncentives = cpVisitsTotal * (Number(s.cpVisitIncentive) || 0) + bookingsCP_total * (Number(s.cpBookingIncentive) || 0);

    const oohItems = [
      { key: 'hoardings', label: `Hoardings (${s.hoardings.count || 0} × ${s.hoardings.months || 0}mo × ₹${(Number(s.hoardings.rate) || 0).toLocaleString('en-IN')})`, value: hoardingsCost },
      { key: 'kiosks', label: 'Wayboards / Kiosks', value: Number(s.kiosks) || 0 },
      { key: 'cpHighTea', label: 'CP High Tea', value: Number(s.cpHighTea) || 0 },
      { key: 'drone', label: 'Drone & Video Shoots', value: (Number(s.droneShoot) || 0) + (Number(s.ugcVideos) || 0) + (Number(s.developerByteShoot) || 0) },
    ];

    const expenses = [
      { key: 'digitalLG', label: 'Digital Lead Gen Spends', value: totalDigitalLGSpendSum, calc: true },
      { key: 'branding', label: 'Branding & Awareness', value: totalBrandingSpendSum, calc: true },
      { key: 'social', label: 'Social Media Boosting', value: Number(s.socialMediaBoosting) || 0 },
      { key: 'drone', label: 'Drone & Construction Location Shoot', value: Number(s.droneShoot) || 0 },
      { key: 'ugc', label: 'UGC Videos — 10 Videos', value: Number(s.ugcVideos) || 0 },
      { key: 'devbyte', label: 'Developer Byte & Testimonial Shoot', value: Number(s.developerByteShoot) || 0 },
      { key: 'chatbot', label: 'Website Chatbot + WhatsApp API Marketing', value: Number(s.chatbotWhatsapp) || 0 },
      { key: 'activation', label: 'Society & Mall Activations', value: Number(s.societyMallActivations) || 0 },
      { key: 'hoardings', label: `Hoardings / OOH (${s.hoardings.count || 0} × ${s.hoardings.months || 0}mo × ₹${(Number(s.hoardings.rate) || 0).toLocaleString('en-IN')})`, value: hoardingsCost },
      { key: 'kiosks', label: 'Kiosks / Wayboards', value: Number(s.kiosks) || 0 },
      { key: 'newspaper', label: 'Newspaper Leaflet Inserts', value: Number(s.newspaperLeaflet) || 0 },
      { key: 'cpgather', label: 'Channel Partner Gathering', value: Number(s.cpGathering) || 0 },
      { key: 'cpincentive', label: 'Channel Partner Incentives', value: cpIncentives, calc: true },
    ];
    const totalAdvertisingBudget = sum(expenses.map(e => e.value));
    const totalRevenue = totalBookings * (Number(s.avgTicketSize) || 0);
    const advertisingCostPct = totalRevenue > 0 ? totalAdvertisingBudget / totalRevenue : 0;
    const cpBrokeragePercent = Number(s.cpBrokeragePercent) || 0;
    const totalMarketingCostPct = advertisingCostPct + cpBrokeragePercent;

    // Channel Partner activation
    const catA = clampArr(s.catA, n);
    const catB = clampArr(s.catB, n);
    const cpDigitalReach = clampArr(s.cpDigitalReach, n);
    const physicalMeetings = months.map((_, i) => ((catA[i] || 0) + (catB[i] || 0)) * 1.8);
    const outreach = physicalMeetings.map(v => v * 4);
    const catA_total = sum(catA), catB_total = sum(catB);
    const physicalMeetings_total = sum(physicalMeetings);
    const outreach_total = sum(outreach);
    const cpDigitalReach_total = sum(cpDigitalReach);

    // Till-date advertising cost %s (Slide 5)
    const td = s.tillDateAdvertising || {};
    function tdCalc(bucket) {
      if (!bucket) return null;
      const totalSpend = (Number(bucket.digital) || 0) + (Number(bucket.ooh) || 0) + (Number(bucket.branding) || 0) + (Number(bucket.shoots) || 0) + (Number(bucket.cpOutgo) || 0);
      const revenue = Number(bucket.revenue) || 0;
      return Object.assign({}, bucket, {
        totalSpend: totalSpend,
        costPct: revenue > 0 ? totalSpend / revenue : 0,
      });
    }
    const tillDate = {
      lifetime: tdCalc(td.lifetime),
      preApril: tdCalc(td.preApril),
      postApril: tdCalc(td.postApril),
    };

    // MTD report totals (Slide 2)
    const mtdMetrics = (s.mtdMetrics || []).map(m => ({
      key: m.key,
      label: m.label,
      budgeted: m.budgeted,
      achieved: m.achieved,
      budgetedTotal: sum(m.budgeted),
      achievedTotal: sum(m.achieved),
    }));

    // Lifetime visits & bookings summary totals (Slide 3)
    const lifetimeHistory = s.lifetimeHistory || [];
    const lifetimeTotals = {
      visits: sum(lifetimeHistory.map(r => r.visits)),
      totalBookings: sum(lifetimeHistory.map(r => r.totalBookings)),
      liveBookings: sum(lifetimeHistory.map(r => r.liveBookings)),
      agreements: sum(lifetimeHistory.map(r => r.agreements)),
    };

    return {
      n: n, months: months,
      bPS: bPS, bDirect: bDirect, bCP: bCP, cpl: cpl, monthBookings: monthBookings,
      visitsPS: visitsPS, visitsDirect: visitsDirect, visitsCP: visitsCP, totalVisits: totalVisits, revisits: revisits,
      presalesSVS: presalesSVS, qualifiedLeads: qualifiedLeads, totalDigitalLeads: totalDigitalLeads,
      digitalLGSpend: digitalLGSpend, brandingSpend: brandingSpend, totalDigitalSpend: totalDigitalSpend, cpqlArr: cpqlArr,
      bookingsPS_total: bookingsPS_total, bookingsDirect_total: bookingsDirect_total, bookingsCP_total: bookingsCP_total, totalBookings: totalBookings,
      totalVisitsSum: totalVisitsSum, totalRevisitsSum: totalRevisitsSum, totalQualifiedLeads: totalQualifiedLeads, totalLeadsSum: totalLeadsSum, totalSVS: totalSVS,
      visitsPS_total: visitsPS_total, visitsDirect_total: visitsDirect_total, visitsCP_total: visitsCP_total,
      convPS: convPS, convDirect: convDirect, convCP: convCP, convOverall: convOverall,
      totalDigitalLGSpendSum: totalDigitalLGSpendSum, totalBrandingSpendSum: totalBrandingSpendSum, totalDigitalSpendSum: totalDigitalSpendSum,
      objectives: objectives, platforms: platforms, campaigns: campaigns, oohItems: oohItems,
      expenses: expenses, totalAdvertisingBudget: totalAdvertisingBudget, totalRevenue: totalRevenue,
      advertisingCostPct: advertisingCostPct, cpBrokeragePercent: cpBrokeragePercent, totalMarketingCostPct: totalMarketingCostPct,
      catA: catA, catB: catB, cpDigitalReach: cpDigitalReach, physicalMeetings: physicalMeetings, outreach: outreach,
      catA_total: catA_total, catB_total: catB_total, physicalMeetings_total: physicalMeetings_total,
      outreach_total: outreach_total, cpDigitalReach_total: cpDigitalReach_total,
      cpVisitsTotal: cpVisitsTotal, cpIncentives: cpIncentives, hoardingsCost: hoardingsCost,
      tillDate: tillDate, mtdMetrics: mtdMetrics, mtdWeeks: s.mtdWeeks || [],
      lifetimeHistory: lifetimeHistory, lifetimeTotals: lifetimeTotals,
      agreementTracker: s.agreementTracker || [],
    };
  }

  /* ---------------------------- Export ---------------------------- */

  global.Calculator = {
    sum, clampArr, round, fmtINR, fmtNum, fmtPct, deriveVisits,
    makeState, SKYLUXE_PRESET, fastTrackPreset, resizeStateToDuration,
    computeModel,
  };

})(window);
