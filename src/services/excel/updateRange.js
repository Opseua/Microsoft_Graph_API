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
    let ret = {
        'ret': false
    };
    retGetRange = await getRange({ sheetTabName: inf.sheetTabName });
    if (!retGetRange.ret) {
        ret['msg'] = `ERRO AO CRIAR SESSAO`;
    } else {
        if (sheetLin == 0) {
            fileId = config.fileId;
            sheetTabName = inf.sheetTabName;
            sheetCol = `A`;
            sheetLin = Number(retGetRange.values);
            token = config.token;
            session = config.session;
        };
        const corpo = { "values": [[`${sheetLin} | ${inf.send}`]] };
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
            ret['msg'] = `ENVIADO→ "${sheetTabName}" ${sheetCol}${sheetLin}`;
            ret['ret'] = true;
            sheetLin += 1;
        }
        else {
            ret['msg'] = `${res.error.message}`;
        }
    }

    console.log(ret.msg);
    return ret
}
export default updateRange



for (let i = 0; i < 1; i++) {
    const ret = await updateRange({ sheetTabName: 'SEAU', send: 'OLÁ' });
    if (ret === false) {
        break;
    }
    await new Promise(resolve => setTimeout(resolve, (1000)));// aguardar 2 segundos
}
