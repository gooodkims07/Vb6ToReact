const http = require('http');

const data = JSON.stringify({
    queryId: 'm_LoginCheck',
    params: {
        sabun: '600018',
        password: 'devpass'
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

        try {
            console.log('Response Body Raw:', responseBody);
            const parsedData = JSON.parse(responseBody);
            if (parsedData.success && parsedData.rows && parsedData.rows.length > 0) {
                console.log('TEST PASSED: Login Successful for 600018');
            } else {
                console.log('TEST FAILED: ' + (parsedData.error || 'Login Failed'));
            }
        } catch (e) {
            console.error('Error parsing JSON:', e);
        }
    });
});

req.on('error', (error) => {
    console.error('Request Error:', error);
});

req.write(data);
req.end();
