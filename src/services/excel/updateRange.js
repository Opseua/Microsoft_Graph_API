import fs from 'fs';
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

let funApi;
async function api(inf) {
    const module = await import('../../resources/api.js');
    funApi = module.default;
    return await funApi(inf);
}
let funCreateSession;
async function createSession(inf) {
    const module = await import('./createSession.js');
    funCreateSession = module.default;
    return await funCreateSession(inf);
}
let funGetRange;
async function getRange(inf) {
    const module = await import('./getRange.js');
    funGetRange = module.default;
    return await funGetRange(inf);
}

let fileId = '';
let retGetRange = '';
let sheetTabName = '';
let sheetCol = '';
let sheetLin = '';
let token = '';
let session = '';

async function updateRange(inf) {
    let msg = '';
    let ret = false;
    const retCreateSession = await createSession();
    if (!retCreateSession) {
        msg = 'ERRO AO CRIAR SESSAO';
    } else {
        if (sheetLin == 0) {
            fileId = config.fileId;
            retGetRange = await getRange();
            sheetTabName = retGetRange.aba;
            sheetCol = `${retGetRange.col}`;
            sheetLin = Number(retGetRange.lin);
            token = config.token;
            session = config.session;
        };
        const corpo = { "values": [[`LINHA: ${sheetLin}`]] };
        const requisicao = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheetTabName}')/range(address='${sheetCol}${sheetLin}')`,
            method: 'PATCH',
            headers: {
                'Content-Type': 'Application/Json',
                'authorization': `Bearer ${token}`,
                'workbook-session-id': `${session}`
            },
            body: corpo
        };
        let res = await api(requisicao);
        res = JSON.parse(res);
        if ("values" in res) {
            msg = `ENVIADO→ ${res.values[0]}`;
            ret = true;
            sheetLin += 1;
        }
        else {
            msg = res.error.message;
        }
    }

    console.log(msg);
    return ret
}
export default updateRange



for (let i = 0; i < 20; i++) {
    const ret = await updateRange();
    if (ret === false) {
        break;
    }
    await new Promise(resolve => setTimeout(resolve, (1000)));// aguardar 2 segundos
}
