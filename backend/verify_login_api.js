const http = require('http');

const data = JSON.stringify({
    queryId: 'm_LoginCheck',
    params: {
        sabun: '600018',
        password: '12345'
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
        console.log('Status:', res.statusCode);
        try {
            const result = JSON.parse(responseBody);
            if (result.success && result.rows && result.rows.length > 0) {
                console.log('RESULT: CORRECT (Login Successful)');
                console.log('User Data:', result.rows[0]);
            } else {
                console.log('RESULT: INCORRECT (Login Failed)');
                console.log('Error/Message:', result.error || 'No user found');
            }
        } catch (e) {
            console.log('Response:', responseBody);
        }
    });
});

req.on('error', (e) => {
    console.error('Request Error:', e);
});

req.write(data);
req.end();
