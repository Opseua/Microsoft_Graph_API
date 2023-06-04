const imp1 = () => import('fs').then(module => module.default);
const fs = await imp1();
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

const api = async (i) => (await import('../../resources/api.js')).default(i);

const createSession = async (i) => (await import('./createSession.js')).default(i);

const getSheetInf = async (i) => (await import('./getSheetInf.js')).default(i);

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
        if ('values' in res) {
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

