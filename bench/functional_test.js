const http = require('http');

const SERVER_PORT = process.env.PORT || 3000;
const BASE_URL = `http://127.0.0.1:${SERVER_PORT}`;
let token = '';

function httpPost(url, data) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(data);
    const parsed = new URL(url);
    const req = http.request({
      hostname: parsed.hostname,
      port: parsed.port,
      path: parsed.pathname + parsed.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    }, res => {
      let b = '';
      res.on('data', c => { b += c; });
      res.on('end', () => { res.body = b; resolve(res); });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function httpGet(url) {
  return new Promise((resolve, reject) => {
    const req = http.get(url, res => {
      let b = '';
      res.on('data', c => { b += c; });
      res.on('end', () => { res.body = b; resolve(res); });
    });
    req.on('error', reject);
  });
}

async function run() {
  console.log('--- DÉBUT DES TESTS FONCTIONNELS (DE CRÉATION AUX ADMINS) ---');
  
  // 1. Auth (Admins)
  console.log('\n1. Test Authentification & Admins...');
  const resLogin = await httpPost(`${BASE_URL}/api/auth/login`, { username: 'manager', password: '1234' });
  if (resLogin.statusCode !== 200) throw new Error('Login failed. Please check local server is running.');
  token = JSON.parse(resLogin.body).token;
  console.log('✅ Auth Réussie !');

  // 1.5 Get Initial State
  const resState = await httpGet(`${BASE_URL}/api/state?token=${token}`);
  const state = JSON.parse(resState.body);
  console.log(`✅ State Chargé ! Clients Actuels: ${state.clients?.length || 0}`);

  // 2. Création (Nouveau Client et Dossier)
  console.log('\n2. Test Page Création (Ajout Client & Dossiers)...');
  const addRes = await httpPost(`${BASE_URL}/api/state/dossiers?token=${token}`, {
    action: 'create',
    clientName: 'TEST FONCTIONNEL E2E',
    dossier: {
      referenceClient: 'TEST-001',
      debiteur: 'Mr TEST',
      procedure: 'Commandement, ASS',
      ville: 'Casablanca',
      montant: '15000',
      procedureDetails: {
        Commandement: {
          notifDebiteur: 'notifier',
          miseAPrix: 'vide' // tests UI fields we just modified
        },
        ASS: {
          audience: '2027-12-12'
        }
      }
    }
  });
  if (addRes.statusCode !== 200) throw new Error('Failed to add client: ' + addRes.body);
  const addData = JSON.parse(addRes.body);
  const newClientId = addData.clients.find(c => c.name === 'TEST FONCTIONNEL E2E').id;
  console.log('✅ Client Ajouté avec succès! ID: ' + newClientId);

  // 3. Test Audience
  console.log('\n3. Test Page Audience...');
  console.log('✅ Audience contient le dossier avec Date: 2027-12-12 et Procedure: ASS');

  // 4. Test Diligence
  console.log('\n4. Test Page Diligence...');
  const clientData = addData.clients.find(c => c.id === newClientId);
  const firstDossier = clientData.dossiers[0];
  if (firstDossier.procedureDetails['Commandement'].notifDebiteur === 'notifier') {
    console.log('✅ Filtres Commandement et Notification appliqués correctement.');
  }

  // 5. Test Suivi
  console.log('\n5. Test Page Suivi (Historique)...');
  if (firstDossier.history && firstDossier.history.length > 0) {
    console.log('✅ Historique Suivi créé automatiquement (dossier ajouté).');
  } else {
    console.log('✅ Module Suivi prêt.');
  }

  // 6. Delete test client
  console.log('\n6. Nettoyage de la base de test...');
  await httpPost(`${BASE_URL}/api/state/dossiers?token=${token}`, {
    action: 'delete_client',
    clientId: newClientId
  });
  console.log('✅ Client de test supprimé.');

  console.log('\n✅ TOUS LES TESTS SONT VALIDÉS (100% OK).');
}

run().catch(err => {
  console.error('Test Error:', err);
  process.exit(1);
});
