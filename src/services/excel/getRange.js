const fs = await import('fs');
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);
const { api } = await import('../../resources/api.js');
const createSession = async (i) => (await import('./createSession.js')).default(i);
const getSheetInf = async (i) => (await import('./getSheetInf.js')).default(i);

async function getRange(inf) {
    let ret = { 'ret': false };
    const retCreateSession = await createSession();
    if (!retCreateSession.ret) {
        ret['msg'] = `ERRO AO CRIAR SESSAO`;
    } else {
        const fileId = config.fileId;
        const retGetSheetInf = await getSheetInf({ sheetTabName: inf.sheetTabName });
        if (!retGetSheetInf.ret) {
            ret['msg'] = `ERRO AO BUSCAR INFORMACOES DA PLANILHA`;
        } else {
            const sheetTabId = retGetSheetInf.res.sheetTabId;
            const token = retGetSheetInf.res.token;
            const session = retCreateSession.res.session;
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
                    ret['res'] = { 'token': token, 'session': session, 'values': res.values[0] };
                    ret['msg'] = `RETORNO: ${res.values[0]}`;
                }
                ret['ret'] = true;
            }
            else {
                ret['msg'] = `${res.error.code}`;
            }
        }
    }
    console.log(ret.msg);
    return ret
}
export default getRange

