const https = require('https');

exports.handler = async function(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

  if (!ANTHROPIC_API_KEY) {
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: 'ERROR: API key no configurada' })
    };
  }

  let body;
  try {
    body = JSON.parse(event.body);
  } catch(e) {
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: 'ERROR: Body invalido - ' + e.message })
    };
  }

  const { prompt } = body;

  const postData = JSON.stringify({
    model: 'claude-sonnet-4-5',
    max_tokens: 1000,
    messages: [{ role: 'user', content: prompt }]
  });

  return new Promise((resolve) => {
    const options = {
      hostname: 'api.anthropic.com',
      path: '/v1/messages',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        resolve({
          statusCode: 200,
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ text: 'RESPUESTA ANTHROPIC: ' + data })
        });
      });
    });

    req.on('error', (e) => {
      resolve({
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text: 'ERROR RED: ' + e.message })
      });
    });

    req.write(postData);
    req.end();
  });
};
