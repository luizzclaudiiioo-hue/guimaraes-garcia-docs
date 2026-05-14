const https = require('https');

module.exports = function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: 'API key not configured' });
    return;
  }

  const bodyStr = JSON.stringify(req.body);

  const options = {
    hostname: 'api.anthropic.com',
    path: '/v1/messages',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Length': Buffer.byteLength(bodyStr),
    },
  };

  const request = https.request(options, function(response) {
    let data = '';
    response.on('data', function(chunk) { data += chunk; });
    response.on('end', function() {
      try {
        res.status(response.statusCode).json(JSON.parse(data));
      } catch (e) {
        res.status(500).json({ error: 'Parse error' });
      }
    });
  });

  request.on('error', function(e) {
    res.status(500).json({ error: e.message });
  });

  request.write(bodyStr);
  request.end();
};
