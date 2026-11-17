"use strict";

/* ==========================================================================
   Dink & Drive – Excel report builder
   Turns one day's records into a workbook, and builds the all-days history.
   (No database access in here – server.js hands in plain data.)
   ========================================================================== */

const ExcelJS = require("exceljs");

const PESO = '"₱"#,##0.00';
const GREEN = "FF14532D";
const LIGHT = "FFE7F5EC";
const GREY = "FFF1F5F9";

const hourLabel = (h) => {
  const period = h >= 12 ? "PM" : "AM";
  const hh = h % 12 === 0 ? 12 : h % 12;
  return `${hh}:00 ${period}`;
};

const sum = (list, fn) => list.reduce((t, x) => t + (Number(fn(x)) || 0), 0);
const money = (n) => Math.round((Number(n) || 0) * 100) / 100;

// ---------------------------------------------------------------------------
// Numbers for one day
// ---------------------------------------------------------------------------
function computeSummary(snap, cfg) {
  const reservations = snap.reservations || [];
  const sales = snap.posSales || [];
  const active = reservations.filter((r) => r.status !== "Cancelled");
  const cancelled = reservations.filter((r) => r.status === "Cancelled");
  const paid = active.filter((r) => r.isPaid);
  const unpaid = active.filter((r) => !r.isPaid);
  const addonAmt = (r) => sum(r.addons || [], (a) => a.amount);

  const courtRevenue = sum(paid, (r) => r.courtFee);
  const addonRevenue = sum(paid, addonAmt);
  const counterRevenue = sum(sales, (s) => s.price);

  // Revenue by payment method
  const byMethod = {};
  for (const r of paid) {
    byMethod[r.paymentStatus] = (byMethod[r.paymentStatus] || 0) + r.total;
  }

  const byChannel = {};
  for (const r of active) {
    const c = (byChannel[r.channel] ||= { bookings: 0, revenue: 0 });
    c.bookings += 1;
    if (r.isPaid) c.revenue += r.total;
  }

  const byCourt = {};
  for (const id of cfg.courtIds) byCourt[id] = { bookings: 0, hours: 0, revenue: 0 };
  for (const r of active) {
    const c = (byCourt[r.court] ||= { bookings: 0, hours: 0, revenue: 0 });
    c.bookings += 1;
    c.hours += r.duration;
    if (r.isPaid) c.revenue += r.courtFee;
  }

  // Items sold: booking add-ons on PAID bookings + counter sales
  const items = {};
  const addItem = (name, qty, amount) => {
    const i = (items[name] ||= { qty: 0, amount: 0 });
    i.qty += qty;
    i.amount += amount;
  };
  for (const r of paid) {
    for (const a of r.addons || []) {
      const m = /^(\d+)x\s+(.+)$/.exec(a.label || "");
      addItem(m ? m[2] : a.label, m ? Number(m[1]) : 1, a.amount);
    }
  }
  for (const s of sales) addItem(s.item, 1, s.price);

  const hoursBooked = sum(active, (r) => r.duration);
  const hoursAvailable =
    cfg.courtIds.length * (cfg.closeHour - cfg.openHour);

  return {
    totalBookings: reservations.length,
    activeBookings: active.length,
    cancelledBookings: cancelled.length,
    completed: active.filter((r) => r.status === "Completed").length,
    noShowOrOpen: active.filter(
      (r) => r.status === "Reserved" || r.status === "Checked-In",
    ).length,
    hoursBooked,
    hoursAvailable,
    courtRevenue: money(courtRevenue),
    addonRevenue: money(addonRevenue),
    bookingRevenue: money(courtRevenue + addonRevenue),
    counterRevenue: money(counterRevenue),
    totalRevenue: money(courtRevenue + addonRevenue + counterRevenue),
    uncollected: money(sum(unpaid, (r) => r.total)),
    uncollectedCount: unpaid.length,
    cancelledPaid: money(
      sum(
        cancelled.filter((r) => r.isPaid),
        (r) => r.total,
      ),
    ),
    cancelledPaidCount: cancelled.filter((r) => r.isPaid).length,
    byMethod,
    byChannel,
    byCourt,
    items,
  };
}

// ---------------------------------------------------------------------------
// Styling helpers
// ---------------------------------------------------------------------------
function styleHeader(row) {
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  row.height = 22;
  row.eachCell((c) => {
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN } };
    c.border = { bottom: { style: "thin", color: { argb: "FF999999" } } };
  });
}

function sectionTitle(ws, text, cols = 3) {
  const row = ws.addRow([text]);
  row.font = { bold: true, size: 12, color: { argb: GREEN } };
  row.eachCell((c, i) => {
    if (i <= cols)
      c.border = { bottom: { style: "medium", color: { argb: GREEN } } };
  });
  for (let i = 2; i <= cols; i++)
    row.getCell(i).border = { bottom: { style: "medium", color: { argb: GREEN } } };
  return row;
}

// ---------------------------------------------------------------------------
// One day's workbook
// ---------------------------------------------------------------------------
async function buildDailyWorkbook(snap, cfg, { partial = false } = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = cfg.clubName;
  wb.created = new Date();
  const s = computeSummary(snap, cfg);

  // ----- Sheet 1: Summary -----
  const ws = wb.addWorksheet("Summary", {
    views: [{ showGridLines: false }],
  });
  ws.columns = [{ width: 38 }, { width: 18 }, { width: 18 }, { width: 18 }];

  const title = ws.addRow([`${cfg.clubName} – Daily Revenue Report`]);
  title.font = { bold: true, size: 16, color: { argb: GREEN } };
  ws.addRow([`Date: ${snap.date}`]).font = { bold: true, size: 12 };
  ws.addRow([
    `Generated: ${snap.generatedAt}${partial ? "   (PARTIAL – the day is not finished yet)" : ""}`,
  ]).font = { italic: true, color: { argb: "FF666666" } };
  ws.addRow([]);

  sectionTitle(ws, "Revenue", 2);
  const rev = [
    ["Court fees (paid bookings)", s.courtRevenue],
    ["Gear / add-ons on bookings (paid)", s.addonRevenue],
    ["Counter sales (POS)", s.counterRevenue],
  ];
  for (const [label, val] of rev) {
    const r = ws.addRow([label, val]);
    r.getCell(2).numFmt = PESO;
  }
  const total = ws.addRow(["TOTAL REVENUE", s.totalRevenue]);
  total.font = { bold: true, size: 12 };
  total.getCell(2).numFmt = PESO;
  total.eachCell((c, i) => {
    if (i <= 2) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: LIGHT } };
  });
  ws.addRow([]);

  sectionTitle(ws, "Money to follow up", 2);
  const u = ws.addRow([
    `Not yet paid (${s.uncollectedCount} booking${s.uncollectedCount === 1 ? "" : "s"})`,
    s.uncollected,
  ]);
  u.getCell(2).numFmt = PESO;
  const cp = ws.addRow([
    `Paid but then cancelled (${s.cancelledPaidCount}) – check refunds`,
    s.cancelledPaid,
  ]);
  cp.getCell(2).numFmt = PESO;
  ws.addRow([]);

  sectionTitle(ws, "Revenue by payment method", 2);
  const methods = Object.entries(s.byMethod);
  if (methods.length === 0) ws.addRow(["(no paid bookings)"]);
  for (const [m, v] of methods) {
    const r = ws.addRow([m, money(v)]);
    r.getCell(2).numFmt = PESO;
  }
  if (s.counterRevenue > 0) {
    const r = ws.addRow(["Counter sales (POS)", s.counterRevenue]);
    r.getCell(2).numFmt = PESO;
  }
  ws.addRow([]);

  sectionTitle(ws, "Bookings", 2);
  ws.addRow(["Total logged", s.totalBookings]);
  ws.addRow(["Active (not cancelled)", s.activeBookings]);
  ws.addRow(["Completed sessions", s.completed]);
  ws.addRow(["Not completed (no-show / still open)", s.noShowOrOpen]);
  ws.addRow(["Cancelled", s.cancelledBookings]);
  const util = s.hoursAvailable
    ? `${s.hoursBooked} of ${s.hoursAvailable} court-hours (${Math.round((s.hoursBooked / s.hoursAvailable) * 100)}%)`
    : `${s.hoursBooked}`;
  ws.addRow(["Court-hours booked", util]);
  ws.addRow([]);

  sectionTitle(ws, "By court", 4);
  styleHeader(ws.addRow(["Court", "Bookings", "Hours", "Court revenue"]));
  for (const [court, v] of Object.entries(s.byCourt)) {
    const r = ws.addRow([court, v.bookings, v.hours, money(v.revenue)]);
    r.getCell(4).numFmt = PESO;
  }
  ws.addRow([]);

  sectionTitle(ws, "By booking channel", 3);
  styleHeader(ws.addRow(["Channel", "Bookings", "Paid revenue"]));
  for (const [ch, v] of Object.entries(s.byChannel)) {
    const r = ws.addRow([ch, v.bookings, money(v.revenue)]);
    r.getCell(3).numFmt = PESO;
  }
  ws.addRow([]);

  sectionTitle(ws, "Items sold (paid bookings + counter)", 3);
  styleHeader(ws.addRow(["Item", "Qty", "Amount"]));
  const items = Object.entries(s.items);
  if (items.length === 0) ws.addRow(["(none)"]);
  for (const [name, v] of items) {
    const r = ws.addRow([name, v.qty, money(v.amount)]);
    r.getCell(3).numFmt = PESO;
  }

  // ----- Sheet 2: Bookings (every transaction) -----
  const bs = wb.addWorksheet("Bookings", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  bs.columns = [
    { header: "Booking ID", key: "id", width: 11 },
    { header: "Receipt No.", key: "receipt", width: 16 },
    { header: "Play Date", key: "date", width: 12 },
    { header: "Channel", key: "channel", width: 13 },
    { header: "Customer", key: "customer", width: 22 },
    { header: "Contact", key: "contact", width: 22 },
    { header: "Court", key: "court", width: 10 },
    { header: "Start", key: "start", width: 10 },
    { header: "End", key: "end", width: 10 },
    { header: "Hours", key: "hours", width: 7 },
    { header: "Court Fee", key: "courtFee", width: 13 },
    { header: "Add-ons", key: "addons", width: 34 },
    { header: "Add-ons Amount", key: "addonAmt", width: 14 },
    { header: "Total", key: "total", width: 13 },
    { header: "Payment", key: "payment", width: 32 },
    { header: "Paid?", key: "paid", width: 8 },
    { header: "Paid At", key: "paidAt", width: 18 },
    { header: "Session Status", key: "status", width: 14 },
    { header: "Processed By", key: "by", width: 20 },
    { header: "Preferences / Notes", key: "notes", width: 34 },
    { header: "Remark", key: "remark", width: 34 },
  ];
  styleHeader(bs.getRow(1));
  const sorted = [...(snap.reservations || [])].sort(
    (a, b) => a.startHour - b.startHour || a.court.localeCompare(b.court),
  );
  for (const r of sorted) {
    let remark = "";
    if (r.status === "Cancelled")
      remark = r.isPaid
        ? "Cancelled AFTER payment – check refund"
        : "Cancelled (no payment taken)";
    else if (!r.isPaid) remark = "Not paid – not counted as revenue";
    const row = bs.addRow({
      id: r.id,
      receipt: r.receiptNo || "",
      date: r.date,
      channel: r.channel,
      customer: r.customer,
      contact: r.contact,
      court: r.court,
      start: hourLabel(r.startHour),
      end: hourLabel(r.startHour + r.duration),
      hours: r.duration,
      courtFee: r.courtFee,
      addons: (r.addons || []).map((a) => a.label).join(", "),
      addonAmt: sum(r.addons || [], (a) => a.amount),
      total: r.total,
      payment: r.paymentStatus,
      paid: r.isPaid ? "Yes" : "No",
      paidAt: r.paidAt || "",
      status: r.status,
      by: r.processedBy || "",
      notes: r.preferences || "",
      remark,
    });
    for (const k of ["courtFee", "addonAmt", "total"]) row.getCell(k).numFmt = PESO;
    if (r.status === "Cancelled")
      row.font = { color: { argb: "FF999999" }, strike: true };
    else if (!r.isPaid) row.font = { color: { argb: "FFB45309" } };
  }
  if (sorted.length) {
    const t = bs.addRow({
      customer: "TOTAL (paid, not cancelled)",
      courtFee: s.courtRevenue,
      addonAmt: s.addonRevenue,
      total: s.bookingRevenue,
    });
    t.font = { bold: true };
    for (const k of ["courtFee", "addonAmt", "total"]) t.getCell(k).numFmt = PESO;
    t.eachCell((c) => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREY } };
    });
    bs.autoFilter = { from: "A1", to: "U1" };
  }

  // ----- Sheet 3: Counter Sales -----
  const ps = wb.addWorksheet("Counter Sales", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  ps.columns = [
    { header: "Sale ID", key: "id", width: 12 },
    { header: "Time", key: "time", width: 20 },
    { header: "Item", key: "item", width: 30 },
    { header: "Price", key: "price", width: 14 },
  ];
  styleHeader(ps.getRow(1));
  for (const x of snap.posSales || []) {
    const row = ps.addRow(x);
    row.getCell("price").numFmt = PESO;
  }
  if ((snap.posSales || []).length) {
    const t = ps.addRow({ item: "TOTAL", price: s.counterRevenue });
    t.font = { bold: true };
    t.getCell("price").numFmt = PESO;
  } else {
    ps.addRow({ item: "(no counter sales logged)" });
  }

  // ----- Sheet 4: Paddle rentals -----
  const ls = wb.addWorksheet("Paddle Rentals", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  ls.columns = [
    { header: "Loan ID", key: "id", width: 11 },
    { header: "Item", key: "item", width: 28 },
    { header: "Holder", key: "holder", width: 30 },
    { header: "Status", key: "status", width: 26 },
  ];
  styleHeader(ls.getRow(1));
  for (const l of snap.loans || []) {
    ls.addRow({
      id: l.id,
      item: l.item,
      holder: l.holder,
      status: l.returned ? `Returned (${l.returnedDay || "?"})` : "NOT RETURNED",
    });
  }
  if (!(snap.loans || []).length) ls.addRow({ item: "(no rentals)" });

  return wb;
}

// ---------------------------------------------------------------------------
// All-days history workbook (one row per day + monthly totals)
// ---------------------------------------------------------------------------
async function buildHistoryWorkbook(rows, cfg) {
  const wb = new ExcelJS.Workbook();
  wb.creator = cfg.clubName;
  wb.created = new Date();
  const sorted = [...rows].sort((a, b) => a.date.localeCompare(b.date));

  const ws = wb.addWorksheet("Daily Totals", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  ws.columns = [
    { header: "Date", key: "date", width: 12 },
    { header: "Bookings", key: "bookings", width: 10 },
    { header: "Cancelled", key: "cancelled", width: 10 },
    { header: "Hours Booked", key: "hours", width: 13 },
    { header: "Court Revenue", key: "court", width: 15 },
    { header: "Add-ons Revenue", key: "addons", width: 16 },
    { header: "Counter Sales", key: "counter", width: 15 },
    { header: "TOTAL REVENUE", key: "total", width: 16 },
    { header: "Cash at Counter", key: "cash", width: 15 },
    { header: "GCash / Transfer", key: "gcash", width: 16 },
    { header: "Not Yet Paid", key: "unpaid", width: 14 },
    { header: "Paid then Cancelled", key: "cancelledPaid", width: 18 },
    { header: "Report File", key: "file", width: 44 },
  ];
  styleHeader(ws.getRow(1));
  const money6 = ["court", "addons", "counter", "total", "cash", "gcash", "unpaid", "cancelledPaid"];
  for (const r of sorted) {
    const row = ws.addRow(r);
    for (const k of money6) row.getCell(k).numFmt = PESO;
    row.getCell("total").font = { bold: true };
    if (r.file)
      row.getCell("file").value = { text: r.file, hyperlink: r.file };
  }
  if (sorted.length) {
    const n = sorted.length + 1;
    const tot = { date: "ALL DAYS" };
    for (const k of ["bookings", "cancelled", "hours", ...money6])
      tot[k] = money(sum(sorted, (r) => r[k]));
    const t = ws.addRow(tot);
    t.font = { bold: true };
    for (const k of money6) t.getCell(k).numFmt = PESO;
    t.eachCell((c) => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: LIGHT } };
    });
    ws.autoFilter = { from: "A1", to: `M${n}` };
  }

  const ms = wb.addWorksheet("Monthly", { views: [{ state: "frozen", ySplit: 1 }] });
  ms.columns = [
    { header: "Month", key: "month", width: 12 },
    { header: "Days Reported", key: "days", width: 14 },
    { header: "Bookings", key: "bookings", width: 10 },
    { header: "Court Revenue", key: "court", width: 15 },
    { header: "Add-ons Revenue", key: "addons", width: 16 },
    { header: "Counter Sales", key: "counter", width: 15 },
    { header: "TOTAL REVENUE", key: "total", width: 16 },
    { header: "Average per Day", key: "avg", width: 16 },
  ];
  styleHeader(ms.getRow(1));
  const months = {};
  for (const r of sorted) (months[r.date.slice(0, 7)] ||= []).push(r);
  for (const [m, list] of Object.entries(months)) {
    const total = sum(list, (r) => r.total);
    const row = ms.addRow({
      month: m,
      days: list.length,
      bookings: sum(list, (r) => r.bookings),
      court: money(sum(list, (r) => r.court)),
      addons: money(sum(list, (r) => r.addons)),
      counter: money(sum(list, (r) => r.counter)),
      total: money(total),
      avg: money(total / list.length),
    });
    for (const k of ["court", "addons", "counter", "total", "avg"])
      row.getCell(k).numFmt = PESO;
  }

  return wb;
}

// Row stored in the history log for one finished day
function historyRow(snap, cfg, file) {
  const s = computeSummary(snap, cfg);
  return {
    date: snap.date,
    bookings: s.activeBookings,
    cancelled: s.cancelledBookings,
    hours: s.hoursBooked,
    court: s.courtRevenue,
    addons: s.addonRevenue,
    counter: s.counterRevenue,
    total: s.totalRevenue,
    cash: money(s.byMethod[cfg.cashMethod] || 0),
    gcash: money(s.byMethod[cfg.gcashMethod] || 0),
    unpaid: s.uncollected,
    cancelledPaid: s.cancelledPaid,
    file,
  };
}

module.exports = {
  computeSummary,
  buildDailyWorkbook,
  buildHistoryWorkbook,
  historyRow,
};
