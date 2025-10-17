function updateAPI() {
  const now = new Date();
  const day = now.getDay(); // Sunday = 0, Monday = 1, ..., Saturday = 6
  const hour = now.getHours(); // 0–23
  const minute = now.getMinutes(); // 0–59

  // Skip after 8:05 PM or on Saturday
  if (hour > 20 || (hour === 20 && minute >= 5) || day === 6) {
    Logger.log("Skipped: Saturday or after 8:05 PM");
    return;
  }

  fetchPvA();
}

/* ALL POSSIBLE COLUMNS — USER CAN REORDER OR REMOVE AS NEEDED */
const ALL_COLUMNS_REFERENCE = [
  'Order URL',
  'Order No',
  'Assigned To',
  'Status',
  'QC Status',
  'QC By',
  'Planned Start',
  'Actual Start',
  'Planned End',
  'Actual End',
  'Location',
  'LDS'
];

function fetchPvA(customFrom = null, customTo = null) {
  const apiKey = '42c853bc30b13910e161834073705caa20qBO5FdcRc';
  const searchUrl = `https://api.optimoroute.com/v1/search_orders?key=${apiKey}`;
  const completionUrl = `https://api.optimoroute.com/v1/get_completion_details?key=${apiKey}`;

  let startDate, endDate;

  if (customFrom && customTo) {
    startDate = customFrom;
    endDate = customTo;
  } else {
    const today = new Date();
    const day = today.getDay();
    const daysSinceSaturday = (day + 1) % 7;
    const lastSaturday = new Date(today);
    lastSaturday.setDate(today.getDate() - daysSinceSaturday);
    const nextFriday = new Date(lastSaturday);
    nextFriday.setDate(lastSaturday.getDate() + 6);
    startDate = lastSaturday.toISOString().slice(0, 10);
    endDate = nextFriday.toISOString().slice(0, 10);
  }

  Logger.log('Importing date range: ' + startDate + ' → ' + endDate);
  Logger.log('Columns being imported: ' + ALL_COLUMNS_REFERENCE.join(', '));

  const baseBody = {
    dateRange: { from: startDate, to: endDate },
    includeOrderData: true,
    includeScheduleInformation: true
  };

  const plannedByOrder = {};
  let afterTag = null;

  while (true) {
    const body = afterTag ? { ...baseBody, after_tag: afterTag } : { ...baseBody };
    const resp = UrlFetchApp.fetch(searchUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) return;

    const json = JSON.parse(resp.getContentText());
    if (!json.success) return;

    const orders = json.orders || [];
    orders.forEach(o => {
      const data = o.data || {};
      const sched = o.scheduleInformation || {};
      const orderNo = data.orderNo;
      if (!orderNo) return;

      plannedByOrder[orderNo] = {
        location: data.location?.locationName || '',
        assignedTo: sched.driverName || '',
        start: sched.scheduledAtDt || '',
        end: sched.scheduledAtDt && typeof data.duration === 'number'
          ? addMinutes(sched.scheduledAtDt, data.duration)
          : '',
        lds: data.customField5 || ''
      };
    });

    afterTag = json.after_tag || null;
    if (!afterTag) break;
  }

  const orderNos = Object.keys(plannedByOrder);
  if (orderNos.length === 0) return;

  const actualByOrder = {};
  const chunkSize = 250;

  for (let i = 0; i < orderNos.length; i += chunkSize) {
    const chunk = orderNos.slice(i, i + chunkSize).map(n => ({ orderNo: n }));
    const body = { orders: chunk };
    const resp = UrlFetchApp.fetch(completionUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) return;

    const json = JSON.parse(resp.getContentText());
    if (!json.success) return;

    (json.orders || []).forEach(r => {
      if (!r || !r.success) return;
      const d = r.data || {};
      actualByOrder[r.orderNo] = {
        start: d.startTime?.localTime || '',
        end: d.endTime?.localTime || '',
        status: d.status || '',
        url: d.tracking_url || ''
      };
    });
  }

  const statusByOrder = fetchStatusesFromSupabase(orderNos);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Google Sheet and run from there');
  let sheet = ss.getSheetByName('PvA API');
  if (!sheet) sheet = ss.insertSheet('PvA API');
  else sheet.clearContents();

  const rows = [ALL_COLUMNS_REFERENCE];
  orderNos.forEach(n => {
    const row = ALL_COLUMNS_REFERENCE.map(col =>
      getColumnValue(col, n, plannedByOrder, actualByOrder, statusByOrder)
    );
    rows.push(row);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  Logger.log('Import complete. Imported ' + (rows.length - 1) + ' orders.');
}

function getColumnValue(column, orderNo, plannedByOrder, actualByOrder, statusByOrder) {
  const p = plannedByOrder[orderNo] || {};
  const a = actualByOrder[orderNo] || {};
  const s = statusByOrder[orderNo] || {};

  switch (column) {
    case 'Order No': return orderNo;
    case 'Assigned To': return p.assignedTo || '';
    case 'Status': return a.status || '';
    case 'QC Status': return s.status || '';
    case 'QC By': return s.last_action_user || '';
    case 'Planned Start': return p.start || '';
    case 'Actual Start': return a.start || '';
    case 'Planned End': return p.end || '';
    case 'Actual End': return a.end || '';
    case 'Location': return p.location || '';
    case 'LDS': return p.lds || '';
    case 'Order URL': return a.url || '';
    default: return '';
  }
}

function addMinutes(dt, minutes) {
  if (!dt || typeof dt !== 'string') return '';
  if (!(typeof minutes === 'number' && isFinite(minutes))) return '';
  const s = dt.indexOf('T') === -1 ? dt.replace(' ', 'T') + 'Z' : dt;
  const d = new Date(s);
  if (isNaN(d.getTime())) return '';
  const d2 = new Date(d.getTime() + minutes * 60000);
  const pad = n => (n < 10 ? '0' + n : '' + n);
  return (
    d2.getUTCFullYear() + '-' +
    pad(d2.getUTCMonth() + 1) + '-' +
    pad(d2.getUTCDate()) + ' ' +
    pad(d2.getUTCHours()) + ':' +
    pad(d2.getUTCMinutes()) + ':' +
    pad(d2.getUTCSeconds())
  );
}

function fetchStatusesFromSupabase(orderNos) {
  const supabaseUrl = 'https://eijdqiyvuhregbydndnb.supabase.co';
  const supabaseKey =
    'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVpamRxaXl2dWhyZWdieWRuZG5iIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTczODA5ODcxMSwiZXhwIjoyMDUzNjc0NzExfQ.pTPdq-7HQuto7T6dgW9dB60hFiMoZgajFCt516tZdl0';
  const table = 'work_orders';
  const statusByOrder = {};
  const chunkSize = 100;

  for (let i = 0; i < orderNos.length; i += chunkSize) {
    const chunk = orderNos.slice(i, i + chunkSize);
    const filter = chunk.join(',');
    const url = `${supabaseUrl}/rest/v1/${table}?select=order_no,status,last_action_user&order_no=in.(${filter})`;

    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        apikey: supabaseKey,
        Authorization: 'Bearer ' + supabaseKey
      },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) continue;

    const rows = JSON.parse(resp.getContentText());
    rows.forEach(r => {
      statusByOrder[r.order_no] = {
        status: r.status,
        last_action_user: r.last_action_user
      };
    });
  }
  return statusByOrder;
}
