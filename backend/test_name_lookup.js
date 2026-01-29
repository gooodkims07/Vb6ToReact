const http = require('http');

const data = JSON.stringify({
    queryId: 'm_GetName',
    params: {
        sabun: '600018'
    }
});

const options = {
    hostname: 'localhost',
    port: 3000,
    path: '/api/query',
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
        'Content-Length': data.length
    }
};

const req = http.request(options, (res) => {
    let responseBody = '';

    res.on('data', (chunk) => {
        responseBody += chunk;
    });

    res.on('end', () => {
        console.log('Response Status:', res.statusCode);
        console.log('Response Body:', responseBody);
    });
});

req.on('error', (error) => {
    console.error('Request Error:', error);
});

req.write(data);
req.end();
