const { normalizeMail, findCodeForMail, findRowByCode, buildProfile } = require('./lib/sheets');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ status: 'error', message: 'POST only' }) };
  }
  try {
    const { mail } = JSON.parse(event.body || '{}');
    const trimmed = normalizeMail(mail);
    if (!trimmed || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
      return json({ status: 'invalid_mail' });
    }

    const code = await findCodeForMail(trimmed);
    if (!code) return json({ status: 'not_registered', mail: trimmed });

    const match = await findRowByCode(code);
    if (!match) return json({ status: 'code_missing_in_sheet', mail: trimmed, code });

    const profile = buildProfile(match.headers, match.row);
    return json({ status: 'ok', profile, code });
  } catch (err) {
    console.error(err);
    return json({ status: 'error', message: err.message }, 500);
  }
};

function json(body, statusCode = 200) {
  return { statusCode, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) };
}
