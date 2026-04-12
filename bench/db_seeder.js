const mysql = require('mysql2/promise');
const path = require('path');
const crypto = require('crypto');
require('dotenv').config({ path: path.join(__dirname, '../server/.env') });

const DB_NAME = 'cabinet_test_db';
const PASSWORD_HASH_ITERATIONS = 120000;

function hashPassword(password) {
    const salt = crypto.randomBytes(16).toString('hex');
    const hash = crypto.pbkdf2Sync(
        password,
        Buffer.from(salt, 'hex'),
        PASSWORD_HASH_ITERATIONS,
        32,
        'sha256'
    ).toString('hex');
    return { salt, hash };
}

async function main() {
    console.log(`Connecting to MySQL to seed ${DB_NAME}...`);
    const connection = await mysql.createConnection({
        host: process.env.DB_HOST || 'localhost',
        port: process.env.DB_PORT || 3306,
        user: process.env.DB_USER || 'root',
        password: process.env.DB_PASSWORD || '',
        database: DB_NAME
    });

    try {
        console.log('Initializing tables...');
        await connection.query(`
            CREATE TABLE IF NOT EXISTS users (
                id BIGINT PRIMARY KEY,
                username VARCHAR(255) UNIQUE,
                passwordHash TEXT,
                passwordSalt TEXT,
                role VARCHAR(50),
                data JSON,
                updatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )
        `);
        await connection.query(`
            CREATE TABLE IF NOT EXISTS clients (
                id BIGINT PRIMARY KEY,
                name VARCHAR(255),
                updatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )
        `);
        await connection.query(`
            CREATE TABLE IF NOT EXISTS dossiers (
                id INT AUTO_INCREMENT PRIMARY KEY,
                clientId BIGINT,
                referenceClient VARCHAR(255),
                debiteur VARCHAR(255),
                procedure_name TEXT,
                data JSON,
                updatedAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                INDEX(clientId),
                INDEX(referenceClient),
                INDEX(debiteur)
            )
        `);

        console.log('Clearing old data...');
        await connection.query('DELETE FROM dossiers');
        await connection.query('DELETE FROM clients');
        await connection.query('DELETE FROM users');

        // 1. Seed Users (11 total)
        console.log('Seeding 11 users...');
        const userDefs = [
            { id: 1, u: 'manager', r: 'manager' },
            { id: 2, u: 'manager2', r: 'manager' },
            { id: 3, u: 'admin1', r: 'admin' },
            { id: 4, u: 'admin2', r: 'admin' },
            { id: 5, u: 'admin3', r: 'admin' },
            { id: 6, u: 'admin4', r: 'admin' },
            { id: 7, u: 'admin5', r: 'admin' },
            { id: 8, u: 'admin6', r: 'admin' },
            { id: 9, u: 'admin7', r: 'admin' },
            { id: 10, u: 'client1', r: 'client' },
            { id: 11, u: 'client2', r: 'client' },
        ];
        
        for (const u of userDefs) {
            const { salt, hash } = hashPassword('1234');
            await connection.query(
                'INSERT INTO users (id, username, passwordHash, passwordSalt, role, data) VALUES (?, ?, ?, ?, ?, ?)',
                [u.id, u.u, hash, salt, u.r, JSON.stringify({ requirePasswordChange: false })]
            );
        }

        // 2. Seed Clients (20 total)
        console.log('Seeding 20 clients...');
        for (let i = 1; i <= 20; i++) {
            await connection.query('INSERT INTO clients (id, name) VALUES (?, ?)', [i, `Client Stress ${i}`]);
        }

        // 3. Seed Dossiers (10,000 total)
        console.log('Seeding 10,000 dossiers (14k audience, 15k diligence)...');
        let audienceCount = 0;
        let diligenceCount = 0;
        const batchSize = 500;

        for (let i = 0; i < 10000; i += batchSize) {
            const batch = [];
            for (let j = 0; j < batchSize; j++) {
                const id = i + j + 1;
                const clientId = (id % 20) + 1;
                const refClient = `R-STRESS-${String(id).padStart(6, '0')}`;
                const debiteur = `Debiteur Stress ${id}`;
                const procedureDetails = {};
                let procedureName = 'ASS';
                
                procedureDetails['ASS'] = { audience: '2026-06-15', juge: 'Juge 1', tribunal: 'Casa', sort: 'En cours' };
                audienceCount++;
                diligenceCount++;

                if (audienceCount < 14000) {
                    procedureDetails['ASS 2'] = { audience: '2026-07-20', juge: 'Juge 2', tribunal: 'Casa', sort: 'En cours' };
                    audienceCount++;
                    procedureName += ', ASS 2';
                }
                if (diligenceCount < 15000) {
                    procedureDetails['SFDC'] = { referenceClient: refClient, ville: 'Casa', statut: 'En cours' };
                    diligenceCount++;
                    procedureName += ', SFDC';
                }

                batch.push([clientId, refClient, debiteur, procedureName, JSON.stringify({
                    id, referenceClient: refClient, debiteur, procedure: procedureName,
                    ville: 'Casablanca', montant: '50000', procedureDetails, archive: false, history: []
                })]);
            }
            await connection.query('INSERT INTO dossiers (clientId, referenceClient, debiteur, procedure_name, data) VALUES ?', [batch]);
            process.stdout.write('.');
        }
        console.log('\nSeed completed successfully!');

    } finally {
        await connection.end();
    }
}

main();
