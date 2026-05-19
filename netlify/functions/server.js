const https = require('https');
const express = require('express');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

app.post('/generar', async (req, res) => {
  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

  if (!ANTHROPIC_API_KEY) {
    return res.json({ error: 'API key no configurada' });
  }

  const { prompt } = req.body;

  const postData = JSON.stringify({
    model: 'claude-sonnet-4-5',
    max_tokens: 2500,
    messages: [{ role: 'user', content: prompt }]
  });

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

  const apiReq = https.request(options, (apiRes) => {
    let data = '';
    apiRes.on('data', (chunk) => { data += chunk; });
    apiRes.on('end', () => {
      try {
        const parsed = JSON.parse(data);
        if (apiRes.statusCode !== 200) {
          res.json({ error: parsed.error?.message || 'Error en la API' });
        } else {
          res.json({ text: parsed.content[0].text });
        }
      } catch(e) {
        res.json({ error: 'Error procesando respuesta: ' + e.message });
      }
    });
  });

  apiReq.on('error', (e) => {
    res.json({ error: e.message });
  });

  apiReq.write(postData);
  apiReq.end();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Flowix API corriendo en puerto ' + PORT));
