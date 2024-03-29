const fs = await import('fs');
const { fileInf } = await import('../../../../Chrome_Extension/src/resources/fileInf.js');
const retfileInf = await fileInf(new URL(import.meta.url).pathname);
const configPath = `${retfileInf.res.pathProject1}\\src\\config.json`
const configFile = fs.readFileSync(configPath);
const config = JSON.parse(configFile);
const { api } = await import('../../../../Chrome_Extension/src/resources/api.js');
const { createSession } = await import('./createSession.js');
const { getSheetInf } = await import('./getSheetInf.js');

async function getRange(inf) {
    let ret = { 'ret': false };
    try {
        const retCreateSession = await createSession();
        if (!retCreateSession.ret) { return ret }

        const fileId = config.fileId;
        const retGetSheetInf = await getSheetInf({ sheetTabName: inf.sheetTabName });
        if (!retGetSheetInf.ret) { return ret }

        const sheetTabId = retGetSheetInf.res.sheetTabId;
        const token = retGetSheetInf.res.token;
        const session = retCreateSession.res.session;
        const infApi = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets/${sheetTabId}/range(address='A1')`,
            method: 'GET',
            headers: {
                'authorization': `Bearer ${token}`,
                'workbook-session-id': `${session}`
            }
        };
        const retApi = await api(infApi);
        if (!retApi.ret) { return ret }

        const res = JSON.parse(retApi.res.body);
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
    } catch (e) {
        ret['msg'] = regexE({ 'e': e }).res;
    }

    if (!ret.ret) { console.log(ret.msg) }
    return ret
}
export { getRange }

