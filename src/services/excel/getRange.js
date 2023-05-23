import fs from 'fs';
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

let fun1;
async function api(inf) {
    const module = await import('../../resources/api.js');
    fun1 = module.default;
    return await fun1(inf);
}
let fun2;
async function createSession(inf) {
    const module = await import('./createSession.js');
    fun2 = module.default;
    return await fun2(inf);
}

const fileId = config.fileId
const sheetTabId = config.sheetTabId
const token = config.token

async function getRange(inf) {
    let session = '';
    if (inf === undefined) {
        session = config.session
    } else {
        session = inf
    }
    const requisicao = {
        url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets/${sheetTabId}/range(address='A2')`,
        method: 'GET',
        headers: {
            'authorization': `Bearer ${token}`,
            'workbook-session-id': `${session}`
        }
    };
    const re = await api(requisicao);
    const res = JSON.parse(re);
    return res;
}

async function run() {
    let msg = '';
    let resultado = await getRange()
    if ("values" in resultado) {
        msg = resultado.values[0];
    }
    else if (resultado.error.code == 'InvalidSession') {
        const session = await createSession();
        resultado = await getRange(session);
        if ("values" in resultado) {
            msg = resultado.values[0];
        } else {
            msg = resultado.error.message;
        }
    } else {
        msg = resultado.error.message;
    }
    console.log(msg)
}
run()