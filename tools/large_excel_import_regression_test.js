const fs = require('fs/promises');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');
const { chromium } = require('playwright');

const SOURCE_ROOT = path.join(__dirname, '..');
const TEMP_ROOT = path.join(os.tmpdir(), `applicationversion1-large-import-${Date.now()}`);
const PORT = 3800;
const HOST = '127.0.0.1';
const BASE_URL = `http://${HOST}:${PORT}`;

const CLIENT_COUNT = 60;
const INITIAL_DOSSIERS = 40000;
const INITIAL_AUDIENCES = 60000;
const EXTRA_DOSSIERS = 3000;
const EXTRA_AUDIENCES = 6000;
const MODIFICATION_COUNT = 120;

async function copyProjectSubset() {
  await fs.mkdir(TEMP_ROOT, { recursive: true });
  for (const entry of [
    'app.js',
    'index.html',
    'style.css',
    'state-persistence.js',
    'render-dashboard.js',
    'render-audience-suivi.js',
    'render-diligence.js',
    'vendor',
    'workers',
    'server'
  ]) {
    await fs.cp(path.join(SOURCE_ROOT, entry), path.join(TEMP_ROOT, entry), { recursive: true });
  }
}

async function writeEmptyServerState() {
  const statePath = path.join(TEMP_ROOT, 'server', 'data', 'state.json');
  await fs.mkdir(path.dirname(statePath), { recursive: true });
  await fs.writeFile(
    statePath,
    JSON.stringify({
      clients: [],
      salleAssignments: [],
      users: [],
      audienceDraft: {},
      recycleBin: [],
      recycleArchive: [],
      version: 0,
      updatedAt: new Date().toISOString()
    }),
    'utf8'
  );
}

async function waitForServer(timeoutMs = 30000) {
  const start = Date.now();
  for (;;) {
    try {
      const res = await fetch(`${BASE_URL}/api/health`);
      if (res.ok) return;
    } catch {}
    if (Date.now() - start > timeoutMs) throw new Error('Server did not become ready in time');
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
}

async function main() {
  await copyProjectSubset();
  await writeEmptyServerState();

  const server = spawn('node', ['server/index.js'], {
    cwd: TEMP_ROOT,
    env: { ...process.env, HOST, PORT: String(PORT) },
    stdio: ['ignore', 'pipe', 'pipe']
  });

  let serverStderr = '';
  server.stderr.on('data', (chunk) => {
    serverStderr += String(chunk);
  });

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  const page = await context.newPage();
  page.setDefaultTimeout(0);

  try {
    await waitForServer();
    await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' });
    await page.fill('#username', 'manager');
    await page.fill('#password', '1234');
    await page.evaluate(() => login());
    await page.waitForSelector('#appContent', { state: 'visible' });

    const result = await page.evaluate(async (config) => {
      const countDossiers = () => (Array.isArray(AppState.clients) ? AppState.clients : []).reduce((sum, client) => (
        sum + (Array.isArray(client?.dossiers) ? client.dossiers.length : 0)
      ), 0);
      const countAudiences = () => getAudienceRows({ ignoreSearch: true, ignoreColor: true }).length;

      const makeClientName = (index) => `Client ${(index % config.clientCount) + 1}`;
      const makeDebiteur = (index) => `Debiteur ${index + 1}`;
      const makeRef = (proc, index) => `${proc === 'ASS' ? 'ASS' : 'RES'}/${index + 1}/2026`;
      const makeProcedureText = () => 'ASS, Restitution';

      function buildDossierRows(startIndex, count) {
        const rows = [];
        for (let i = 0; i < count; i += 1) {
          const globalIndex = startIndex + i;
          rows.push({
            rowNumber: globalIndex + 2,
            clientName: makeClientName(globalIndex),
            dateAffectation: `${String((globalIndex % 28) + 1).padStart(2, '0')}/${String(((globalIndex + 4) % 12) + 1).padStart(2, '0')}/2025`,
            carriedAffectationDate: '',
            carriedMontant: '',
            dateAffectationExtra: '',
            type: ['Auto', 'Credit', 'Recouvrement'][globalIndex % 3],
            procedureText: makeProcedureText(),
            refClient: makeRef('ASS', globalIndex),
            debiteur: makeDebiteur(globalIndex),
            montant: String(1000 + (globalIndex % 90000)),
            montantExtra: '',
            immatriculation: `WW-${100000 + globalIndex}`,
            boiteNo: '',
            caution: '',
            marque: ['Dacia', 'Renault', 'Peugeot', 'Hyundai'][globalIndex % 4],
            adresse: `Adresse ${globalIndex + 1}`,
            ville: ['Casablanca', 'Rabat', 'Fes', 'Agadir'][globalIndex % 4],
            cautionAdresse: '',
            cautionVille: '',
            cautionCin: '',
            cautionRc: '',
            refAssignation: makeRef('ASS', globalIndex),
            refRestitution: makeRef('Restitution', globalIndex),
            refSfdc: '',
            refInjonction: '',
            notificationSort: '',
            notificationNo: '',
            executionNo: '',
            sort: '',
            tribunal: '',
            statutRaw: globalIndex % 5 === 0 ? 'Suspension / test' : ''
          });
        }
        return rows;
      }

      function buildAudienceRows(startAudienceIndex, count, totalDossiers) {
        const rows = [];
        for (let i = 0; i < count; i += 1) {
          const audienceIndex = startAudienceIndex + i;
          const dossierIndex = audienceIndex % totalDossiers;
          const proc = audienceIndex < totalDossiers ? 'ASS' : 'Restitution';
          rows.push({
            rowNumber: audienceIndex + 2,
            refClient: makeRef('ASS', dossierIndex),
            debiteur: makeDebiteur(dossierIndex),
            refDossier: makeRef(proc, dossierIndex),
            procedureText: proc,
            audience: `${String((audienceIndex % 28) + 1).padStart(2, '0')}/${String(((audienceIndex + 1) % 12) + 1).padStart(2, '0')}/2026`,
            juge: `Juge ${audienceIndex % 80}`,
            sort: `Sort ${audienceIndex % 7}`,
            tribunal: ['Casablanca', 'Rabat', 'Fes', 'Agadir'][audienceIndex % 4],
            dateDepot: `${String((audienceIndex % 28) + 1).padStart(2, '0')}/${String(((audienceIndex + 3) % 12) + 1).padStart(2, '0')}/2025`
          });
        }
        return rows;
      }

      const initialPayload = {
        dossiers: buildDossierRows(0, config.initialDossiers),
        audiences: buildAudienceRows(0, config.initialAudiences, config.initialDossiers),
        referenceHints: {}
      };
      await applyExcelImport(initialPayload, { importDossiers: true, importAudiences: true });
      const afterInitial = {
        clients: AppState.clients.length,
        dossiers: countDossiers(),
        audiences: countAudiences()
      };

      const dossierOnlyPayload = {
        dossiers: buildDossierRows(config.initialDossiers, config.extraDossiers),
        audiences: [],
        referenceHints: {}
      };
      await applyExcelImport(dossierOnlyPayload, {
        importDossiers: true,
        importAudiences: false,
        clearAudienceOnDossierOnly: false
      });
      const afterDossierOnly = {
        clients: AppState.clients.length,
        dossiers: countDossiers(),
        audiences: countAudiences()
      };

      const audienceOnlyPayload = {
        dossiers: [],
        audiences: buildAudienceRows(config.initialAudiences, config.extraAudiences, config.initialDossiers + config.extraDossiers),
        referenceHints: {}
      };
      await applyExcelImport(audienceOnlyPayload, {
        importDossiers: false,
        importAudiences: true,
        audienceMode: 'audience-only'
      });
      const afterAudienceOnly = {
        clients: AppState.clients.length,
        dossiers: countDossiers(),
        audiences: countAudiences()
      };

      const flatTargets = [];
      AppState.clients.forEach((client, clientIndex) => {
        (client.dossiers || []).forEach((dossier, dossierIndex) => {
          flatTargets.push({ clientId: Number(client.id), clientIndex, dossierIndex, ref: String(dossier?.referenceClient || '') });
        });
      });

      const touchedRefs = [];
      const stride = Math.max(1, Math.floor(flatTargets.length / config.modificationCount));
      for (let i = 0; i < flatTargets.length && touchedRefs.length < config.modificationCount; i += stride) {
        const target = flatTargets[i];
        const client = AppState.clients[target.clientIndex];
        const dossier = client?.dossiers?.[target.dossierIndex];
        if (!client || !dossier) continue;
        const updated = {
          ...dossier,
          note: `bulk-note-${touchedRefs.length}`,
          avancement: `bulk-step-${touchedRefs.length}`,
          statut: touchedRefs.length % 2 === 0 ? 'Clôture' : 'Suspension'
        };
        client.dossiers[target.dossierIndex] = updated;
        await persistDossierPatchNow({
          action: 'update',
          clientId: Number(client.id),
          dossierIndex: target.dossierIndex,
          targetClientId: Number(client.id),
          dossier: updated
        }, { source: 'large-import-regression' });
        touchedRefs.push({
          clientId: Number(client.id),
          referenceClient: String(updated.referenceClient || ''),
          note: updated.note,
          avancement: updated.avancement,
          statut: updated.statut
        });
      }

      const remoteState = await fetch(`${API_BASE}/state`, { cache: 'no-store' }).then((res) => res.json());
      const remoteClients = Array.isArray(remoteState?.clients) ? remoteState.clients : [];
      let preservedModificationCount = 0;
      touchedRefs.forEach((target) => {
        const client = remoteClients.find((item) => Number(item?.id) === Number(target.clientId));
        const dossier = (client?.dossiers || []).find((item) => String(item?.referenceClient || '') === target.referenceClient);
        if (
          dossier
          && String(dossier.note || '') === target.note
          && String(dossier.avancement || '') === target.avancement
          && String(dossier.statut || '') === target.statut
        ) {
          preservedModificationCount += 1;
        }
      });

      return {
        afterInitial,
        afterDossierOnly,
        afterAudienceOnly,
        checks: {
          audiencesPreservedAfterDossierOnly: afterDossierOnly.audiences >= afterInitial.audiences,
          audiencesIncreasedAfterAudienceOnly: afterAudienceOnly.audiences > afterDossierOnly.audiences,
          preservedModificationCount,
          expectedModificationCount: touchedRefs.length
        }
      };
    }, {
      clientCount: CLIENT_COUNT,
      initialDossiers: INITIAL_DOSSIERS,
      initialAudiences: INITIAL_AUDIENCES,
      extraDossiers: EXTRA_DOSSIERS,
      extraAudiences: EXTRA_AUDIENCES,
      modificationCount: MODIFICATION_COUNT
    });

    console.log(JSON.stringify(result, null, 2));
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
    server.kill('SIGINT');
    await new Promise((resolve) => setTimeout(resolve, 500));
    if (server.exitCode === null) server.kill('SIGKILL');
    if (serverStderr.trim()) console.error(serverStderr.trim());
  }
}

if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
