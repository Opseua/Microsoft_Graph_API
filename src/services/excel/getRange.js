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

async function getRange(inf) {
    let msg = '';
    let ret = false;
    const retCreateSession = await createSession();
    if (!retCreateSession) {
        msg = 'ERRO AO CRIAR SESSAO';
    } else {
        const fileId = config.fileId;
        const sheetTabIdINF = config.sheetTabIdINF;
        const token = config.token;
        const session = config.session;
        const requisicao = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets/${sheetTabIdINF}/range(address='A2')`,
            method: 'GET',
            headers: {
                'authorization': `Bearer ${token}`,
                'workbook-session-id': `${session}`
            }
        };
        let res = await api(requisicao);
        res = JSON.parse(res);
        if ("values" in res) {
            if (JSON.stringify(res.values[0]).includes('\\":\\"')) {
                msg = JSON.parse(`{${res.values[0]}  "x":"x"}`);
                msg = {
                    aba: msg.aba,
                    col: `${msg.range_ini}`,
                    lin: `${msg.lin_vazia}`
                };;
            } else {
                msg = `RETORNO: ${res.values[0]}`;
            }
            //ret = true;
            ret = msg;
        }
        else {
            msg = res.error.code;
        }
    }

    console.log(msg);
    return ret
}
export default getRange

