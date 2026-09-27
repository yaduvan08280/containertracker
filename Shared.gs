/**
 * Container Tracker — shared definitions.
 * Everything in this file is ALSO sent to the browser (see sharedJs_()),
 * so the pipeline, fields and deadline maths live in exactly one place.
 * Keep it plain ES5-ish JS with no Apps Script services.
 */

var DEF = (function () {
  var D = {
    app: 'Container Tracker',
    locations: ['JPR/JPR', 'JPR/MUN', 'DHNK/MUN', 'MUN/MUN', 'AHM/AHM', 'DADRI/JPR', 'SNPT/JPT', 'SNPT/MUN'],

    // Main pipeline — one per container, strictly sequential.
    steps: [
      { n: 1, key: 'S1', name: 'Booking Details', short: 'Booking', fields: ['bookingNo', 'line', 'vessel', 'party', 'destination', 'numContainers'] },
      { n: 2, key: 'S2', name: 'Container Allotment Date', short: 'Allot Date', fields: ['allotDate'], bulk: true },
      { n: 3, key: 'S3', name: 'Container Allotment Details', short: 'Allotment', fields: ['arrivalDateFactory', 'containerNo', 'invoiceNo'] },
      { n: 4, key: 'S4', name: 'S.I & Gate-In Cut-offs', short: 'Cut-offs', fields: ['siCutoff', 'location', 'gateInCutoff'], bulk: true },
      { n: 5, key: 'S5', name: 'S.I Submission', short: 'SI Submit', fields: ['siRemarks'], mark: true },
      { n: 6, key: 'S6', name: 'Container Departure Date (Factory)', short: 'Depart Date', fields: ['departDate'], bulk: true },
      { n: 7, key: 'S7', name: 'Stuffing / Container Dispatch from Factory', short: 'Dispatch', fields: ['estReachDate', 'driverContact', 'vehicleNo', 'dispatchRemark'] },
      { n: 8, key: 'S8', name: 'All Clearance Documents Preparation', short: 'Docs Prep', fields: ['docPrepRemark'], mark: true },
      { n: 9, key: 'S9', name: 'All Clearance Documents Submission', short: 'Docs Submit', fields: ['docSubRemark'], mark: true },
      { n: 10, key: 'S10', name: 'Container Gate-In', short: 'Gate-In', fields: ['gateInDate', 'gateInRemark'] }
    ],

    // Side sections — maintained until the admin closes the container file.
    sections: [
      { key: 'TRACK', name: 'Shipment Tracking', hint: 'ETD / BL / ETA / Arrival / ISF / DO / Clearance',
        fields: ['etd', 'blDate', 'eta', 'arrivalDate', 'isf', 'filingDate', 'doNo', 'clearance', 'arrivalNotice'] },
      { key: 'DOCS', name: 'Shipping Bill & BL', hint: 'Mandatory once the container is dispatched',
        fields: ['sbNo', 'sbPhoto', 'blNo', 'blPhoto'] },
      { key: 'CHARGES', name: 'Charges & Expenses', hint: 'Detention / extra costs with proof photos', fields: [] }
    ],

    chargeGroups: [
      { key: 'TR', name: 'Transportation Charges', items: [['trExtra', 'Extra Transportation Charge']] },
      { key: 'SL', name: 'Shipping Line Charges', items: [
        ['slPrivateYard', 'Private Yard Charges'], ['slRollover', 'Rollover'], ['slPickup', 'Container Pickup Charge'],
        ['slDetention', 'Detention'], ['slSiAmend', 'S.I Amendment'], ['slLateSi', 'Late S.I. Submit'],
        ['slGri', 'GRI / Freight Diff'], ['slSsr', 'SSR Charge'], ['slVesselChange', 'Vessel Change Charge']] },
      { key: 'EX', name: 'Examination Charges', items: [['exCfs', 'CFS Charge'], ['exReseal', 'Re-Seal Charge'], ['exCustom', 'Custom Charges']] },
      { key: 'AG', name: 'After Gate-In', items: [['agGroundRent', 'Ground Rent'], ['agDemurrage', 'Demurrage Charges'], ['agExam', 'Examination Charges']] },
      { key: 'DS', name: 'Destination Charges', items: [['dsDemurrage', 'Demurrage'], ['dsReturnDet', 'Return Detention']] }
    ],
    extraSlots: 5,

    // t: text | int | money | date | datetime | select | textarea | tel | file
    fields: {
      bookingNo: { l: 'Booking No.', t: 'text', req: 1, up: 1 },
      line: { l: 'Line', t: 'text', req: 1 },
      vessel: { l: 'Vessel', t: 'text', req: 1 },
      party: { l: 'Party Name', t: 'text', req: 1 },
      destination: { l: 'Destination', t: 'text', req: 1 },
      numContainers: { l: 'No. of Containers', t: 'int', req: 1, ro: 1 },

      allotDate: { l: 'Container Allotment Date', t: 'date', req: 1 },
      arrivalDateFactory: { l: 'Container Arrival Date (Factory)', t: 'date', req: 1 },
      containerNo: { l: 'Container No.', t: 'text', req: 1, up: 1 },
      invoiceNo: { l: 'Invoice No.', t: 'text', req: 1 },
      siCutoff: { l: 'S.I Cut-off Date', t: 'datetime', req: 1 },
      location: { l: 'Receiving / Handover Location', t: 'select', req: 1 },
      gateInCutoff: { l: 'Gate-In Cut-off Date', t: 'datetime', req: 1 },
      siRemarks: { l: 'Remarks', t: 'textarea' },
      departDate: { l: 'Container Departure Date (Factory)', t: 'datetime', req: 1 },
      estReachDate: { l: 'Estimated Reaching Date (Port/Station)', t: 'datetime', req: 1 },
      driverContact: { l: 'Driver Contact No.', t: 'tel', req: 1 },
      vehicleNo: { l: 'Vehicle No.', t: 'text', req: 1, up: 1 },
      dispatchRemark: { l: 'Remark', t: 'textarea' },
      docPrepRemark: { l: 'Remarks', t: 'textarea' },
      docSubRemark: { l: 'Remarks', t: 'textarea' },
      gateInDate: { l: 'Actual Gate-In Date', t: 'datetime', req: 1 },
      gateInRemark: { l: 'Remarks', t: 'textarea' },

      etd: { l: 'ETD', t: 'date' },
      blDate: { l: 'Actual BL Date', t: 'date' },
      eta: { l: 'ETA', t: 'date' },
      arrivalDate: { l: 'Actual Arrival Date', t: 'date' },
      isf: { l: 'ISF', t: 'text' },
      filingDate: { l: 'Filing Date', t: 'date' },
      doNo: { l: 'DO', t: 'text' },
      clearance: { l: 'Clearance', t: 'text' },
      arrivalNotice: { l: 'Arrival Notice', t: 'text' },

      sbNo: { l: 'Shipping Bill No.', t: 'text', req: 1, up: 1 },
      sbPhoto: { l: 'Shipping Bill Photo', t: 'file', req: 1 },
      blNo: { l: 'BL No.', t: 'text', req: 1, up: 1 },
      blPhoto: { l: 'BL Photo', t: 'file', req: 1 }
    },

    // Fields that can drive a deadline (admin picks one per step).
    deadlineFields: ['allotDate', 'arrivalDateFactory', 'siCutoff', 'gateInCutoff', 'departDate', 'estReachDate'],

    // Default deadline rules. src:'after' = hours after previous step completed;
    // src:'field' = that date field (+/- hours offset). DOCS = hours after dispatch (step 7).
    defaultRules: {
      S2: { src: 'after', hours: 24 },
      S3: { src: 'field', field: 'allotDate', hours: 0 },
      S4: { src: 'after', hours: 24 },
      S5: { src: 'field', field: 'siCutoff', hours: 0 },
      S6: { src: 'after', hours: 24 },
      S7: { src: 'field', field: 'departDate', hours: 0 },
      S8: { src: 'after', hours: 2 },
      S9: { src: 'after', hours: 2 },
      S10: { src: 'field', field: 'gateInCutoff', hours: 0 },
      DOCS: { src: 'after', hours: 24 }
    },

    roles: { admin: 'Admin', pc: 'Process Coordinator', user: 'User' }
  };

  // Attach owner (step / section key) to every field, and generate charge fields.
  D.steps.forEach(function (s) { s.fields.forEach(function (k) { D.fields[k].own = s.key; }); });
  D.sections.forEach(function (s) { s.fields.forEach(function (k) { D.fields[k].own = s.key; }); });
  D.chargeGroups.forEach(function (g) {
    g.items.forEach(function (it) {
      D.fields[it[0]] = { l: it[1], t: 'money', own: 'CHARGES', grp: g.key };
      D.fields[it[0] + '_p1'] = { l: it[1] + ' — Photo 1', t: 'file', own: 'CHARGES', grp: g.key };
      D.fields[it[0] + '_p2'] = { l: it[1] + ' — Photo 2', t: 'file', own: 'CHARGES', grp: g.key };
    });
  });
  for (var i = 1; i <= D.extraSlots; i++) {
    D.fields['ex' + i + '_desc'] = { l: 'Extra Expense ' + i + ' — Description', t: 'text', own: 'CHARGES', grp: 'XX' };
    D.fields['ex' + i + '_amt'] = { l: 'Extra Expense ' + i + ' — Amount', t: 'money', own: 'CHARGES', grp: 'XX' };
    D.fields['ex' + i + '_p1'] = { l: 'Extra Expense ' + i + ' — Photo 1', t: 'file', own: 'CHARGES', grp: 'XX' };
    D.fields['ex' + i + '_p2'] = { l: 'Extra Expense ' + i + ' — Photo 2', t: 'file', own: 'CHARGES', grp: 'XX' };
  }
  D.stepKeys = D.steps.map(function (s) { return s.key; }).concat(D.sections.map(function (s) { return s.key; }));
  return D;
})();

/** Deadline (ms) of pipeline step n for container c, or null if it can't be known yet. */
function ctDue(c, n, rules) {
  var r = rules && rules['S' + n];
  if (!r) return null;
  var h = (Number(r.hours) || 0) * 3600000;
  if (r.src === 'field') {
    var v = c[r.field];
    if (v === undefined || v === null || v === '') return null;
    var t = Number(v);
    var f = DEF.fields[r.field];
    if (f && f.t === 'date') { var d = new Date(t); d.setHours(23, 59, 0, 0); t = d.getTime(); }
    return t + h;
  }
  var p = c['s' + (n - 1) + '_at'];
  return p ? Number(p) + h : null;
}

/** Derived state of a container: current step, deadline, overdue flags, late count. */
function ctState(c, rules, now) {
  var cur = 0, done = 0, late = 0, lateSteps = [];
  for (var n = 1; n <= 10; n++) {
    var at = c['s' + n + '_at'];
    if (at) {
      done++;
      var du = c['s' + n + '_due'];
      if (du && Number(at) > Number(du)) { late++; lateSteps.push(n); }
    } else if (!cur) cur = n;
  }
  var active = (c.status || 'ACTIVE') === 'ACTIVE';
  var due = active && cur ? ctDue(c, cur, rules) : null;
  var docsNeeded = active && !!c.s7_at && !(c.sbNo && c.sbPhoto && c.blNo && c.blPhoto);
  var dr = (rules && rules.DOCS) || {};
  var docsDue = docsNeeded ? Number(c.s7_at) + (Number(dr.hours) || 24) * 3600000 : null;
  return {
    cur: active ? cur : 0, done: done, due: due, over: !!(due && now > due),
    docsNeeded: docsNeeded, docsDue: docsDue, docsOver: !!(docsDue && now > docsDue),
    late: late, lateSteps: lateSteps, active: active
  };
}
