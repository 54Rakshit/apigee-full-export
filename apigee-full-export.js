const axios = require('axios');
const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

const ORG = 'apigee-migration-485807';
const TOKEN = execSync('gcloud auth print-access-token').toString().trim();
const headers = { Authorization: `Bearer ${TOKEN}` };

async function get(url) {
    console.log('GET:', url);
    const res = await axios.get(url, { headers });
    return res.data;
}

// ---------- STEP 1: PROXIES ----------

async function getProxies() {
    const data = await get(`https://apigee.googleapis.com/v1/organizations/${ORG}/apis`);
    console.log('APIs found:', data.proxies.length);
    return data.proxies.map(p => p.name);
}

async function getLatestRevision(api) {
    const revs = await get(`https://apigee.googleapis.com/v1/organizations/${ORG}/apis/${api}/revisions`);
    const latest = revs[revs.length - 1];
    console.log(`Latest rev for ${api}:`, latest);
    return latest;
}

async function extractBundleInfo(api) {
  const rev = await getLatestRevision(api);

  const res = await axios.get(
    `https://apigee.googleapis.com/v1/organizations/${ORG}/apis/${api}/revisions/${rev}?format=bundle`,
    { headers, responseType: 'arraybuffer' }
  );

  const zip = new AdmZip(res.data);

  let basePath = '';
  let targetUrl = '';
  let policies = [];
  let policyContents = [];

  for (const entry of zip.getEntries()) {
    const name = entry.entryName;

    if (name.startsWith('apiproxy/proxies/')) {
      const xml = zip.readAsText(entry);
      const m = xml.match(/<BasePath>(.*?)<\/BasePath>/);
      if (m) basePath = m[1];
    }

    if (name.startsWith('apiproxy/targets/')) {
      const xml = zip.readAsText(entry);
      const m = xml.match(/<URL>(.*?)<\/URL>/);
      if (m) targetUrl = m[1];
    }

    // ✅ Policy names + FULL XML
    if (name.startsWith('apiproxy/policies/')) {
      const fileName = name.split('/').pop();
      const xml = zip.readAsText(entry);

      policies.push(fileName);
      policyContents.push(`===== ${fileName} =====\n${xml}`);
    }
  }

  return {
    basePath,
    targetUrl,
    policies: policies.join(', '),
    policyXML: policyContents.join('\n\n')
  };
}


// ---------- STEP 2: PRODUCTS ----------

async function buildProductMap() {
    const map = {};
    const data = await get(`https://apigee.googleapis.com/v1/organizations/${ORG}/apiproducts`);
    console.log('Products:', data.apiProduct.length);

    for (const p of data.apiProduct) {
        const name = typeof p === 'string' ? p : p.name;
        const detail = await get(`https://apigee.googleapis.com/v1/organizations/${ORG}/apiproducts/${name}`);

        console.log(`Product ${name} apiResources:`, detail.apiResources);

        const ops = detail.operationGroup?.operationConfigs || [];

        for (const op of ops) {
            const apiSource = op.apiSource;

            if (!map[apiSource]) map[apiSource] = [];
            map[apiSource].push(name);
        }
    }

    console.log('Final productMap:', map);
    return map;
}

// ---------- STEP 3: DEVELOPERS ----------

async function getDevelopers() {
    const d = await get(`https://apigee.googleapis.com/v1/organizations/${ORG}/developers`);
    console.log('Developers:', d.developer.length);
    return d.developer;
}

async function getApps(email) {
    const d = await get(`https://apigee.googleapis.com/v1/organizations/${ORG}/developers/${email}/apps`);
    console.log(`Apps for ${email}:`, d.app);
    return d.app;
}

async function getAppDetail(email, app) {
    console.log(`Getting app detail: ${email} -> ${app}`);
    return await get(`https://apigee.googleapis.com/v1/organizations/${ORG}/developers/${email}/apps/${app}`);
}

// ---------- MAIN ----------

(async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Apigee Export');

    sheet.columns = [
        { header: 'API Proxy', key: 'api', width: 25 },
        { header: 'BasePath', key: 'basePath', width: 20 },
        { header: 'Target URL', key: 'targetUrl', width: 40 },
        { header: 'Policies', key: 'policies', width: 40 },
        { header: 'Policy XML', key: 'policyXML', width: 100 },
        { header: 'Product', key: 'product', width: 25 },
        { header: 'Developer', key: 'dev', width: 30 },
        { header: 'App Name', key: 'app', width: 25 },
        { header: 'API Key', key: 'key', width: 40 },
    ];

    console.log('==== STEP 1: PROXIES ====');
    const apis = await getProxies();

    const apiInfo = {};
    for (const api of apis) {
        apiInfo[api] = await extractBundleInfo(api);
    }

    console.log('==== STEP 2: PRODUCTS ====');
    const productMap = await buildProductMap();

    console.log('==== STEP 3: DEVELOPERS ====');
    const developers = await getDevelopers();

    for (const dev of developers) {
        const email = dev.email;
        const apps = await getApps(email);

        for (const a of apps) {

            const appId = a.appId;

            const detail = await getAppDetail(email, appId);

            const appName = detail.name;

            for (const cred of detail.credentials) {
                const key = cred.consumerKey;

                const rawProducts =
                    Array.isArray(cred.apiProducts)
                        ? cred.apiProducts
                        : cred.apiProducts?.apiProduct || [];

                for (const p of rawProducts) {

                    const product =
                        typeof p === 'string'
                            ? p
                            : p.apiproduct; 

                    console.log('Checking product', product, 'for key', key);

                    for (const api in apiInfo) {

                        if (productMap[api] && productMap[api].includes(product)) {

                            console.log('MATCH FOUND:', { api, product });

                            sheet.addRow({
                                api,
                                ...apiInfo[api],
                                product,
                                dev: email,
                                app: appName,
                                key
                            });
                        }

                    }

                }

            }
        }
    }
    // Add APIs that are NOT part of any product (no API keys)
    console.log('==== STEP 4: ADDING APIS WITH NO PRODUCT ====');

    for (const api in apiInfo) {

        // check if this api already exists in sheet
        const alreadyAdded = sheet.getRows(2, sheet.rowCount)
            ?.some(r => r.getCell('api').value === api);

        if (!alreadyAdded) {

            console.log('Adding API with NO PRODUCT:', api);

            sheet.addRow({
                api,
                ...apiInfo[api],
                product: 'NO PRODUCT (No API Key)',
                dev: '',
                app: '',
                key: ''
            });
        }
    }


    await workbook.xlsx.writeFile('apigee-bundle-export.xlsx');
    console.log('------- DONE — Excel created');
})();
