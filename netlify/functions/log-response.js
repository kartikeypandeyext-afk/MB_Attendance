const { logResponse } = require('./lib/sheets');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ status: 'error', message: 'POST only' }) };
  }
  try {
    const payload = JSON.parse(event.body || '{}');
    await logResponse(payload);
    return { statusCode: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ status: 'ok' }) };
  } catch (err) {
    console.error(err);
    // Never let a logging failure surface as an error to the user-facing UI
    return { statusCode: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ status: 'error', message: err.message }) };
  }
};
