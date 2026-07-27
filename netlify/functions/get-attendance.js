const {
  TIMEZONE, findRowByCode, buildProfile, getDateColumns, getPlannedForCode, classifyStatus
} = require('./lib/sheets');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ status: 'error', message: 'POST only' }) };
  }
  try {
    const { code } = JSON.parse(event.body || '{}');
    const match = await findRowByCode(code);
    if (!match) return json({ status: 'invalid_code' });

    const { headers, row } = match;
    const dateCols = getDateColumns(headers);
    const now = new Date();
    const today = new Date(now); today.setHours(0, 0, 0, 0);
    const yesterday = new Date(today); yesterday.setDate(yesterday.getDate() - 1);

    const days = [];
    for (const col of dateCols) {
      if (col.date > today) continue; // hide future dates, same as Code.gs
      const raw = String(row[col.index] || '').trim();
      const statusInfo = classifyStatus(raw, col.date, today, now);
      days.push({
        dateLabel: col.label,
        isoDate: isoDate(col.date),
        weekday: col.date.toLocaleDateString('en-US', { weekday: 'short', timeZone: TIMEZONE }),
        isToday: col.date.getTime() === today.getTime(),
        isYesterday: col.date.getTime() === yesterday.getTime(),
        status: statusInfo
      });
    }
    days.sort((a, b) => new Date(b.isoDate) - new Date(a.isoDate));

    const plannedToday = await getPlannedForCode(code, today);

    return json({ status: 'ok', profile: buildProfile(headers, row), days, plannedToday });
  } catch (err) {
    console.error(err);
    return json({ status: 'error', message: err.message }, 500);
  }
};

function isoDate(d) {
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
function json(body, statusCode = 200) {
  return { statusCode, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) };
}
