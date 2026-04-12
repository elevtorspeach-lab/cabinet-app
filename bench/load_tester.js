const http = require('http');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../server/.env') });

const BASE_URL = 'http://localhost:3000';
const USERS = [
    { username: 'manager', password: '1234' },
    { username: 'manager2', password: '1234' },
    { username: 'admin1', password: '1234' },
    { username: 'admin2', password: '1234' },
    { username: 'admin3', password: '1234' },
    { username: 'admin4', password: '1234' },
    { username: 'admin5', password: '1234' },
    { username: 'admin6', password: '1234' },
    { username: 'admin7', password: '1234' },
    { username: 'client1', password: '1234' },
    { username: 'client2', password: '1234' },
];

async function post(url, data, token = null) {
    return new Promise((resolve, reject) => {
        const body = JSON.stringify(data);
        const parsed = new URL(url);
        const options = {
            hostname: parsed.hostname,
            port: parsed.port,
            path: parsed.pathname + (token ? `?token=${token}` : ''),
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(body)
            }
        };
        const req = http.request(options, (res) => {
            let resBody = '';
            res.on('data', (d) => resBody += d);
            res.on('end', () => resolve({ statusCode: res.statusCode, body: resBody }));
        });
        req.on('error', reject);
        req.write(body);
        req.end();
    });
}

async function get(url, token = null) {
    return new Promise((resolve, reject) => {
        const parsed = new URL(url);
        const fullPath = parsed.pathname + (token ? `?token=${token}` : '') + (parsed.search ? (token ? '&' : '?') + parsed.search.slice(1) : '');
        const options = {
            hostname: parsed.hostname,
            port: parsed.port,
            path: fullPath,
            method: 'GET'
        };
        const req = http.get(options, (res) => {
            let resBody = '';
            res.on('data', (d) => resBody += d);
            res.on('end', () => resolve({ statusCode: res.statusCode, body: resBody }));
        });
        req.on('error', reject);
    });
}

async function simulateUser(user) {
    console.log(`[User ${user.username}] Starting simulation...`);
    try {
        // 1. Login
        const loginRes = await post(`${BASE_URL}/api/auth/login`, { username: user.username, password: user.password });
        if (loginRes.statusCode !== 200) throw new Error(`Login failed for ${user.username}`);
        const { token } = JSON.parse(loginRes.body);
        console.log(`[User ${user.username}] Authenticated.`);

        // 2. Create Dosiers (repeat 5 times)
        for(let i = 0; i < 5; i++) {
            const start = Date.now();
            const dossierRes = await post(`${BASE_URL}/api/state/dossiers`, {
                action: 'create',
                clientId: 1,
                dossier: {
                    referenceClient: `LOAD-TEST-${user.username}-${Date.now()}-${i}`,
                    debiteur: `Load Test Debiteur ${i}`,
                    procedure: 'ASS',
                    ville: 'Casablanca',
                    montant: '1000'
                }
            }, token);
            const latency = Date.now() - start;
            if (dossierRes.statusCode === 200) {
                console.log(`[User ${user.username}] Created dossier ${i+1} (${latency}ms)`);
            } else {
                console.error(`[User ${user.username}] Failed to create dossier ${i+1}: ${dossierRes.statusCode}`);
            }
            await new Promise(r => setTimeout(r, 500)); // typing delay
        }

        return true;
    } catch (err) {
        console.error(`[User ${user.username}] Error: ${err.message}`);
        return false;
    }
}

async function runExport(userId, token) {
    console.log(`[Export Worker] Starting export for user ${userId}...`);
    const start = Date.now();
    const res = await get(`${BASE_URL}/api/state/export-page?page=0&limit=100`, token);
    const latency = Date.now() - start;
    if (res.statusCode === 200) {
        console.log(`[Export Worker] Export success for user ${userId} (${latency}ms)`);
    } else {
        console.error(`[Export Worker] Export failed for user ${userId}: ${res.statusCode}`);
    }
}

async function main() {
    console.log('--- STARTING LOAD TEST ---');
    console.log(`Targeting ${BASE_URL}`);

    // Login as manager to get tokens for export simulation
    const mgrLogin = await post(`${BASE_URL}/api/auth/login`, { username: 'manager', password: '1234' });
    const mgrToken = JSON.parse(mgrLogin.body).token;
    const mgr2Login = await post(`${BASE_URL}/api/auth/login`, { username: 'manager2', password: '1234' });
    const mgr2Token = JSON.parse(mgr2Login.body).token;

    console.log('Simulating 11 concurrent users creating dossiers + 2 concurrent exports...');
    
    const startTime = Date.now();
    
    const userPromises = USERS.map(user => simulateUser(user));
    const exportPromises = [
        runExport('manager', mgrToken),
        runExport('manager2', mgr2Token)
    ];

    const results = await Promise.all([...userPromises, ...exportPromises]);
    
    const duration = Date.now() - startTime;
    console.log('--- TEST FINISHED ---');
    console.log(`Total duration: ${(duration / 1000).toFixed(2)}s`);
    
    const successUsers = results.slice(0, 11).filter(r => r === true).length;
    console.log(`Success Rate (Users): ${successUsers}/11`);
}

main();
