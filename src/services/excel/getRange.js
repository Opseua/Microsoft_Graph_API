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
let fun3;
async function getSheetInf(inf) {
    const module = await import('./getSheetInf.js');
    fun3 = module.default;
    return await fun3(inf);
}

async function getRange(inf) {
    let ret = { 'ret': false };
    const retCreateSession = await createSession();
    if (!retCreateSession) {
        ret['msg'] = `ERRO AO CRIAR SESSAO`;
    } else {
        const fileId = config.fileId;
        const retGetSheetInf = await getSheetInf({ sheetTabName: inf.sheetTabName });
        const sheetTabId = retGetSheetInf.sheetTabId;
        const token = config.token;
        const session = config.session;
        const requisicao = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets/${sheetTabId}/range(address='A1')`,
            method: 'GET',
            headers: {
                'authorization': `Bearer ${token}`,
                'workbook-session-id': `${session}`
            }
        };
        const retApi = await api(requisicao);
        const res = JSON.parse(retApi.res);
        if ("values" in res) {
            if (JSON.stringify(res.values[0]).includes('\\":\\"')) {
                msg = JSON.parse(`{${res.values[0]}  "x":"x"}`);
                msg = {
                    aba: msg.aba,
                    col: `${msg.range_ini}`,
                    lin: `${msg.lin_vazia}`
                };
                ret['msg'] = `${msg}`;
            } else {
                ret['values'] = `${res.values[0]}`;
                ret['msg'] = `RETORNO: ${res.values[0]}`;
            }
            ret['ret'] = true;
        }
        else {
            ret['msg'] = `${res.error.code}`;
        }
    }

    console.log(ret.msg);
    return ret
}
export default getRange
