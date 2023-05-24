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

let celula = 'TESTE!A';
let numero = 10;

async function updateRange(inf) {
    let msg = '';
    let ret = false;
    const retcreateSession = await createSession();
    if (!retcreateSession) {
        let msg = 'ERRO AO CRIAR SESSAO';
    } else {
        const fileId = config.fileId;
        const sheetTabName = config.sheetTabName;
        const sheetRange = config.sheetRange;
        const token = config.token;
        const session = config.session;
        numero += 1;
        const corpo = { "values": [[`${celula}${numero}`]] };
        const requisicao = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheetTabName}')/range(address='${celula}${numero}')`,
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
            msg = res.values[0];
            ret = true;
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
    await new Promise(resolve => setTimeout(resolve, (2000)));// aguardar 2 segundos
}
