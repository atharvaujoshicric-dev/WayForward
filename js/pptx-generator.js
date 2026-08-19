/* ==========================================================================
   pptx-generator.js
   Builds the full 17-slide widescreen PowerPoint deck with PptxGenJS,
   mirroring the interactive slide viewer in app.js.
   ========================================================================== */

(function (global) {
  'use strict';

  const C = global.Calculator;

  const NAVY = '0F172A', SLATE = '1E293B', SLATE2 = '334155', GOLD = 'D97706', AMBER = 'F59E0B',
    WHITE = 'FFFFFF', LIGHT = 'F8FAFC', GRAY = '64748B', LINE = 'E2E8F0';

  function headerRow(cells) {
    return cells.map(c => ({ text: c, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 9.5 } }));
  }

  function tblOpts(extra) {
    return Object.assign({ x: 0.5, w: 12.33, fontSize: 9.5, border: { type: 'solid', color: LINE, pt: 0.5 }, autoPage: false, valign: 'middle' }, extra || {});
  }

  function titleBlock(slide, eyebrow, title, dark) {
    slide.addText(eyebrow.toUpperCase(), { x: 0.5, y: 0.32, w: 10, h: 0.3, fontSize: 10.5, color: GOLD, bold: true, charSpacing: 2 });
    slide.addText(title, { x: 0.5, y: 0.58, w: 12.3, h: 0.65, fontSize: 23, bold: true, color: dark ? WHITE : NAVY, fontFace: 'Georgia' });
    slide.addShape('rect', { x: 0.5, y: 1.2, w: 1.3, h: 0.04, fill: { color: GOLD } });
  }

  function addStatChips(slide, y, items) {
    const w = (12.33 - (items.length - 1) * 0.2) / items.length;
    items.forEach((it, i) => {
      const x = 0.5 + i * (w + 0.2);
      slide.addShape('roundRect', { x, y, w, h: 0.95, fill: { color: LIGHT }, line: { color: LINE, width: 0.75 }, rectRadius: 0.06 });
      slide.addText(it[1], { x, y: y + 0.06, w, h: 0.5, fontSize: 17, bold: true, color: GOLD, align: 'center' });
      slide.addText(it[0], { x, y: y + 0.55, w, h: 0.32, fontSize: 8.5, color: GRAY, align: 'center' });
    });
  }

  function addStatChipsDark(slide, x, y, w, items) {
    items.forEach((it, i) => {
      const yy = y + i * 1.1;
      slide.addShape('roundRect', { x, y: yy, w, h: 0.92, fill: { color: SLATE }, line: { color: GOLD, width: 0.75 }, rectRadius: 0.06 });
      slide.addText(it[1], { x, y: yy + 0.06, w, h: 0.48, fontSize: 15, bold: true, color: GOLD, align: 'center' });
      slide.addText(it[0], { x, y: yy + 0.52, w, h: 0.32, fontSize: 8.5, color: 'CBD5E1', align: 'center' });
    });
  }

  async function generate(state, model) {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
    pptx.layout = 'WIDE';

    const s = state, m = model;

    /* SLIDE 1 — COVER */
    let s1 = pptx.addSlide(); s1.background = { color: NAVY };
    s1.addShape('rect', { x: 0, y: 7.35, w: 13.33, h: 0.06, fill: { color: GOLD } });
    s1.addText(`${s.projectName} WAY FORWARD PLAN`, { x: 0.8, y: 2.5, w: 11.7, h: 1.1, fontSize: 42, bold: true, color: WHITE, fontFace: 'Georgia' });
    s1.addShape('rect', { x: 0.8, y: 3.75, w: 1.4, h: 0.05, fill: { color: GOLD } });
    s1.addText(`${s.months[0]} – ${s.months[m.n - 1]} ${s.year}`, { x: 0.8, y: 3.95, w: 10, h: 0.5, fontSize: 16, color: GOLD, bold: true });
    s1.addText([
      { text: `₹${(m.totalRevenue / 1e7).toFixed(2)} Cr `, options: { bold: true, color: GOLD } },
      { text: 'Revenue Target   •   ', options: { color: 'CBD5E1' } },
      { text: `${m.totalBookings} `, options: { bold: true, color: GOLD } },
      { text: 'Bookings   •   ', options: { color: 'CBD5E1' } },
      { text: `${C.fmtNum(m.totalVisitsSum)} `, options: { bold: true, color: GOLD } },
      { text: 'Site Visits', options: { color: 'CBD5E1' } },
    ], { x: 0.8, y: 4.6, w: 11, h: 0.5, fontSize: 13 });

    /* SLIDE 2 — MTD REPORT */
    let s2 = pptx.addSlide(); s2.background = { color: WHITE };
    titleBlock(s2, 'Section 01', `${s.months[0]} Report — MTD Tracking`, false);
    const mtdHead = ['Metric', ...s.mtdWeeks.flatMap(w => [w + ' Bud.', w + ' Ach.'])];
    const mtdRows = m.mtdMetrics.map(mm => [mm.label, ...mm.budgeted.flatMap((b, i) => [C.fmtNum(b), C.fmtNum(mm.achieved[i])])]);
    s2.addTable([headerRow(mtdHead), ...mtdRows], tblOpts({ y: 1.5, fontSize: 8.5, colW: [2.2, ...Array(8).fill(1.27)] }));

    /* SLIDE 3 — LIFETIME VISITS & BOOKING SUMMARY */
    let s3 = pptx.addSlide(); s3.background = { color: WHITE };
    titleBlock(s3, 'Section 02', 'Lifetime Visits & Booking Summary — Month-wise', false);
    const lifeRows = m.lifetimeHistory.map(r => [r.month, C.fmtNum(r.visits), C.fmtNum(r.totalBookings), C.fmtNum(r.liveBookings), C.fmtNum(r.agreements)]);
    s3.addTable([headerRow(['Month', 'Visits', 'Total Bookings', 'Live Bookings', 'Agreements']), ...lifeRows], tblOpts({ y: 1.5, w: 8, colW: [2, 1.5, 1.5, 1.5, 1.5] }));
    addStatChipsDark(s3, 9, 1.5, 3.83, [
      ['Total Visits', C.fmtNum(m.lifetimeTotals.visits)],
      ['Total Bookings', C.fmtNum(m.lifetimeTotals.totalBookings)],
      ['Agreements Done', C.fmtNum(m.lifetimeTotals.agreements)],
    ]);

    /* SLIDE 4 — AGREEMENT TRACKER */
    let s4 = pptx.addSlide(); s4.background = { color: NAVY };
    titleBlock(s4, 'Section 03', 'Agreement Tracker with Remarks', true);
    const agRows = m.agreementTracker.map(r => [String(r.sr), r.client, r.unit, r.status]);
    s4.addTable([headerRow(['Sr No.', 'Client Name', 'Unit Number', 'Agreement Status & Remarks'])].concat(agRows),
      tblOpts({ y: 1.5, color: WHITE, fill: { color: SLATE }, border: { type: 'solid', color: SLATE2, pt: 0.5 }, colW: [0.8, 2.8, 1.8, 6.93] }));

    /* SLIDE 5 — TILL-DATE ADVERTISING COST */
    let s5 = pptx.addSlide(); s5.background = { color: WHITE };
    titleBlock(s5, 'Section 04', 'Till-Date Advertising Cost Calculation', false);
    const td = m.tillDate;
    const rowsDef = [
      ['Revenue', b => C.fmtINR(b.revenue, true)], ['Digital', b => C.fmtINR(b.digital, true)],
      ['OOH', b => C.fmtINR(b.ooh, true)], ['Branding', b => C.fmtINR(b.branding, true)],
      ['Shoots', b => C.fmtINR(b.shoots, true)], ['CP Outgo', b => C.fmtINR(b.cpOutgo, true)],
      ['Total Marketing Spend', b => C.fmtINR(b.totalSpend, true)], ['Advertising Cost %', b => C.fmtPct(b.costPct)],
    ];
    const tdRows = rowsDef.map(([label, fn]) => [label, fn(td.lifetime), fn(td.preApril), fn(td.postApril)]);
    s5.addTable([headerRow(['Metric', 'Lifetime', 'Pre-April', 'Post-April']), ...tdRows], tblOpts({ y: 1.5, w: 8, colW: [3, 1.67, 1.67, 1.67] }));

    /* SLIDE 6 — DIVIDER */
    let s6 = pptx.addSlide(); s6.background = { color: NAVY };
    s6.addText('SECTION 05', { x: 0, y: 3.0, w: 13.33, h: 0.4, fontSize: 13, color: GOLD, bold: true, align: 'center', charSpacing: 3 });
    s6.addText('Way Forward Strategy', { x: 0, y: 3.4, w: 13.33, h: 1.0, fontSize: 40, bold: true, color: WHITE, fontFace: 'Georgia', align: 'center' });
    s6.addShape('rect', { x: 5.9, y: 4.5, w: 1.5, h: 0.04, fill: { color: GOLD } });

    /* SLIDE 7 — STRATEGIC TARGETS & CHANNEL ATTRIBUTION */
    let s7 = pptx.addSlide(); s7.background = { color: WHITE };
    titleBlock(s7, 'Section 06', 'Strategic Targets & Channel Attribution', false);
    addStatChips(s7, 1.5, [
      ['Duration', m.n + ' Months'], ['Total Bookings', String(m.totalBookings)],
      ['Digital + Direct', String(m.bookingsPS_total + m.bookingsDirect_total)], ['Channel Partner', String(m.bookingsCP_total)],
    ]);
    const channelRows = [
      ['Digital / Presales', C.fmtNum(m.totalLeadsSum), C.fmtNum(m.visitsPS_total), String(m.bookingsPS_total)],
      ['Direct / Walk-in', '—', C.fmtNum(m.visitsDirect_total), String(m.bookingsDirect_total)],
      ['Channel Partner', '—', C.fmtNum(m.visitsCP_total), String(m.bookingsCP_total)],
    ];
    s7.addTable([headerRow(['Tool', 'Leads', 'Estimated Walk-ins', 'Bookings']), ...channelRows], tblOpts({ y: 2.7, w: 9 }));

    /* SLIDE 8 — MONTH-WISE FUNNEL TARGETS (90 DAYS) */
    let s8 = pptx.addSlide(); s8.background = { color: NAVY };
    titleBlock(s8, 'Section 07', 'Month-Wise Funnel Targets (90 Days)', true);
    const funnelRows = m.months.map((mo, i) => [
      mo, C.fmtINR(m.totalDigitalSpend[i], true), C.fmtNum(m.totalDigitalLeads[i]), C.fmtINR(m.cpl[i]),
      C.fmtNum(m.qualifiedLeads[i]), C.fmtINR(m.cpqlArr[i]), C.fmtNum(m.presalesSVS[i]), C.fmtNum(m.visitsPS[i]),
      C.fmtNum(m.visitsDirect[i]), C.fmtNum(m.visitsCP[i]), C.fmtNum(m.revisits[i]), String(m.monthBookings[i]),
    ]);
    s8.addTable([headerRow(['Month', 'Spend', 'Leads', 'CPL', 'Qual. Leads', 'CPQL', 'SVS', 'PS Visits', 'Walk-ins', 'CP Visits', 'Revisits', 'Bookings']), ...funnelRows],
      tblOpts({ y: 1.5, fontSize: 8.5, color: WHITE, fill: { color: SLATE }, border: { type: 'solid', color: SLATE2, pt: 0.5 } }));

    /* SLIDE 9 — DIGITAL FUNNEL & OBJECTIVE/PLATFORM SPLIT */
    let s9 = pptx.addSlide(); s9.background = { color: WHITE };
    titleBlock(s9, 'Section 08', 'Digital Funnel Breakdown & Allocation', false);
    s9.addTable([
      headerRow(['Stage', 'Value']),
      ['Total Digital Leads', C.fmtNum(m.totalLeadsSum)],
      ['Qualified Leads', C.fmtNum(m.totalQualifiedLeads)],
      ['Site Visits (Presales)', C.fmtNum(m.visitsPS_total)],
      ['Bookings (Presales)', String(m.bookingsPS_total)],
    ], tblOpts({ y: 1.55, w: 5.7, colW: [3.6, 2.1] }));
    const objRows = Object.values(m.objectives).map(o => [o.label, C.fmtPct(o.leadPct), C.fmtNum(o.leads), C.fmtPct(o.budgetPct), C.fmtINR(o.budget, true)]);
    s9.addTable([headerRow(['Objective', 'Leads %', 'Leads', 'Budget %', 'Budget']), ...objRows], tblOpts({ x: 6.5, y: 1.55, w: 6.33, colW: [2.0, 1.0, 1.13, 1.0, 1.2] }));
    const platRows = m.platforms.map(p => [p.label, p.pct ? C.fmtPct(p.pct) : '—', p.cpl ? C.fmtINR(p.cpl) : '—', C.fmtNum(p.leads), p.budget ? C.fmtINR(p.budget, true) : '₹0']);
    s9.addTable([headerRow(['Platform', 'Weight', 'Bench. CPL', 'Leads', 'Budget']), ...platRows], tblOpts({ x: 6.5, y: 4.0, w: 6.33, colW: [2.0, 1.0, 1.13, 1.0, 1.2] }));

    /* SLIDE 10 — DIGITAL DETAILED BIFURCATION (CAMPAIGN MATRIX) */
    let s10 = pptx.addSlide(); s10.background = { color: NAVY };
    titleBlock(s10, 'Section 09', 'Digital Detailed Bifurcation — Campaign Matrix', true);
    const campRows = m.campaigns.map(c => [c.name, c.objective, c.audience, C.fmtINR(c.budget, true), C.fmtPct(c.pct)]);
    s10.addTable([headerRow(['Campaign', 'Objective', 'Audience Targeting', 'Monthly Budget', '% Budget']), ...campRows],
      tblOpts({ y: 1.5, fontSize: 9, color: WHITE, fill: { color: SLATE }, border: { type: 'solid', color: SLATE2, pt: 0.5 }, colW: [3.2, 1.9, 4.03, 1.7, 1.5] }));

    /* SLIDE 11 — AUDIENCE TARGETING & MESSAGING FRAMEWORK */
    let s11 = pptx.addSlide(); s11.background = { color: WHITE };
    titleBlock(s11, 'Section 10', 'Audience Targeting & Messaging Framework', false);
    const tiers = [
      { tag: '1ST TOUCH — LEAD GENERATION', rows: [['Demographics', 'HNI & upgrade buyers, 32–55'], ['Target Geos', 'City core, premium suburbs, NRI corridors'], ['Mediums', 'Meta, Google Search, Pmax'], ['Comm. Angle', 'Aspiration + scarcity'], ['Content', 'Walkthroughs, amenity films']] },
      { tag: '2ND TOUCH — REMARKETING', rows: [['Demographics', 'Engaged leads & website visitors'], ['Target Geos', 'Retarget pools from 1st-touch'], ['Mediums', 'Meta remarketing, Display, WhatsApp'], ['Comm. Angle', 'Urgency + social proof'], ['Content', 'Testimonials, offer carousels']] },
    ];
    tiers.forEach((t, i) => {
      const x = 0.5 + i * 6.2;
      s11.addShape('roundRect', { x, y: 1.55, w: 5.9, h: 5.2, fill: { color: LIGHT }, line: { color: LINE, width: 0.75 }, rectRadius: 0.06 });
      s11.addText(t.tag, { x: x + 0.25, y: 1.8, w: 5.4, h: 0.35, fontSize: 11, color: GOLD, bold: true, charSpacing: 1.5 });
      let ty = 2.35;
      t.rows.forEach(([label, val]) => {
        s11.addText(label, { x: x + 0.25, y: ty, w: 5.4, h: 0.28, fontSize: 10, bold: true, color: NAVY });
        s11.addText(val, { x: x + 0.25, y: ty + 0.28, w: 5.4, h: 0.5, fontSize: 10, color: GRAY });
        ty += 0.92;
      });
    });

    /* SLIDE 12 — OOH & OFFLINE ACTIVATION SPLIT */
    let s12 = pptx.addSlide(); s12.background = { color: NAVY };
    titleBlock(s12, 'Section 11', 'OOH & Offline Activation Split', true);
    const oohRows = m.oohItems.map(o => [o.label, C.fmtINR(o.value)]);
    s12.addTable([headerRow(['Marketing Tools', 'Cost']), ...oohRows], tblOpts({ y: 1.5, w: 8, color: WHITE, fill: { color: SLATE }, border: { type: 'solid', color: SLATE2, pt: 0.5 }, colW: [6, 2] }));
    addStatChips(s12, 5.9, [['Total Offline Investment', C.fmtINR(m.oohItems.reduce((a, o) => a + o.value, 0), true)]]);

    /* SLIDE 13 — CHANNEL PARTNER ACTIVATION TARGETS */
    let s13 = pptx.addSlide(); s13.background = { color: WHITE };
    titleBlock(s13, 'Section 12', 'Channel Partner Activation Targets', false);
    const cpRows = m.months.map((mo, i) => [mo, String(m.cpDigitalReach[i]), C.fmtNum(m.outreach[i]), C.fmtNum(m.physicalMeetings[i]), String(m.catA[i]), String(m.catB[i]), C.fmtNum(m.visitsCP[i]), String(m.bCP[i])]);
    s13.addTable([headerRow(['Month', 'Digital Reach', 'WA/Email Outreach', 'Physical Meetings', 'CAT A', 'CAT B', 'CP Visits', 'Conversions']), ...cpRows], tblOpts({ y: 1.5, fontSize: 9 }));

    /* SLIDE 14 — CREATIVE SHOWCASE */
    let s14 = pptx.addSlide(); s14.background = { color: NAVY };
    titleBlock(s14, 'Section 13', 'Creative Showcase — Campaign Formats', true);
    ['Location Advantage', 'Lifestyle Amenities', 'Skyline & Interiors', 'Developer Trust Films'].forEach((label, i) => {
      const x = 0.5 + (i % 2) * 6.2, y = 1.6 + Math.floor(i / 2) * 2.7;
      s14.addShape('roundRect', { x, y, w: 5.9, h: 2.5, fill: { color: SLATE }, line: { color: SLATE2, width: 0.75 }, rectRadius: 0.06 });
      s14.addShape('rect', { x: x + 0.3, y: y + 0.3, w: 5.3, h: 1.5, fill: { color: '0B1220' } });
      s14.addText(label, { x: x + 0.3, y: y + 1.9, w: 5.3, h: 0.4, fontSize: 12, bold: true, color: WHITE });
    });

    /* SLIDE 15 — FESTIVE OFFERS */
    let s15 = pptx.addSlide(); s15.background = { color: WHITE };
    titleBlock(s15, 'Section 14', 'Festive Offers & CP Referral Schemes', false);
    ['Independence Day Offer', 'Festive Season Bonanza', 'Early-Bird Booking Benefits', 'CP Referral Scheme'].forEach((label, i) => {
      const x = 0.5 + (i % 2) * 6.2, y = 1.6 + Math.floor(i / 2) * 2.7;
      s15.addShape('roundRect', { x, y, w: 5.9, h: 2.5, fill: { color: LIGHT }, line: { color: GOLD, width: 1 }, rectRadius: 0.06 });
      s15.addShape('rect', { x: x + 0.3, y: y + 0.3, w: 5.3, h: 1.5, fill: { color: 'FDE9CC' } });
      s15.addText(label, { x: x + 0.3, y: y + 1.9, w: 5.3, h: 0.4, fontSize: 12, bold: true, color: NAVY });
    });

    /* SLIDE 16 — 90-DAY BUDGET BIFURCATION */
    let s16 = pptx.addSlide(); s16.background = { color: NAVY };
    titleBlock(s16, 'Section 15', 'Comprehensive 90-Day Media Plan & Cost Bifurcation', true);
    const expRows = m.expenses.map(e => [e.label, C.fmtINR(e.value)]);
    s16.addTable([headerRow(['Line Item', 'Investment']), ...expRows], tblOpts({ y: 1.5, w: 8.2, color: WHITE, fill: { color: SLATE }, border: { type: 'solid', color: SLATE2, pt: 0.5 }, colW: [6.2, 2] }));
    addStatChipsDark(s16, 9, 1.5, 3.83, [
      ['Total Advertising Budget', C.fmtINR(m.totalAdvertisingBudget, true)],
      ['Total Revenue', C.fmtINR(m.totalRevenue, true)],
      ['Advertising Cost %', C.fmtPct(m.advertisingCostPct)],
      ['Net Marketing Cost %', C.fmtPct(m.totalMarketingCostPct)],
    ]);

    /* SLIDE 17 — CLOSING */
    let s17 = pptx.addSlide(); s17.background = { color: NAVY };
    s17.addText('THANK YOU', { x: 0, y: 2.9, w: 13.33, h: 0.4, fontSize: 13, color: GOLD, bold: true, align: 'center', charSpacing: 3 });
    s17.addText(s.projectName, { x: 0, y: 3.3, w: 13.33, h: 1.0, fontSize: 38, bold: true, color: WHITE, fontFace: 'Georgia', align: 'center' });
    s17.addShape('rect', { x: 5.9, y: 4.35, w: 1.5, h: 0.04, fill: { color: GOLD } });
    s17.addText('For queries, reach the marketing team.', { x: 0, y: 4.6, w: 13.33, h: 0.35, fontSize: 12, color: 'CBD5E1', align: 'center' });
    s17.addText('[ Contact placeholder ]  ·  [ Email placeholder ]  ·  [ Phone placeholder ]', { x: 0, y: 4.95, w: 13.33, h: 0.35, fontSize: 10, color: '94A3B8', align: 'center' });

    await pptx.writeFile({ fileName: `${s.projectName.replace(/\s+/g, '_')}_Way_Forward_Plan.pptx` });
  }

  global.PptxGenerator = { generate };

})(window);
