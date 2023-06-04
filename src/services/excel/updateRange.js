const imp1 = () => import('fs').then(module => module.default);
const fs = await imp1();
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

const api = async (i) => (await import('../../resources/api.js')).default(i);

const getRange = async (i) => (await import('./getRange.js')).default(i);

let fileId, retGetRange, sheetTabName, sheetCol, sheetLin, token, session, run = 0

async function updateRange(inf) {
    let ret = { 'ret': false };
    if (run == 0) {
        retGetRange = await getRange({ sheetTabName: inf.sheetTabName });
        if (!retGetRange.ret) {
            run = -1;
            console.log('GET RANGE: ERRO');
        } else {
            run = 1;
            fileId = config.fileId;
            sheetTabName = inf.sheetTabName;
            sheetCol = `A`;
            sheetLin = Number(retGetRange.values);
            token = config.token;
            session = config.session;
        }
    }
    if (run == 1) {
        const requisicao = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheetTabName}')/range(address='${sheetCol}${sheetLin}')`,
            method: 'PATCH',
            headers: {
                'Content-Type': 'Application/Json',
                'authorization': `Bearer ${token}`,
                'workbook-session-id': `${session}`
            },
            body: { 'values': [[`${sheetLin} | ${inf.send}`]] }
        };
        const retApi = await api(requisicao);
        const res = JSON.parse(retApi.res);
        if (!('values' in res)) {
            run = -1;
            ret['msg'] = `${res.error.message}`;
        } else {
            ret['msg'] = `ENVIADO→ "${sheetTabName}" ${sheetCol}${sheetLin}`;
            ret['ret'] = true;
            sheetLin += 1;
        }
    }

    console.log(ret.msg);
    return ret
}
export default updateRange



// for (let i = 0; i < 5; i++) {
//     const ret = await updateRange({ sheetTabName: 'HAUPC', send: 'OLÁ' });
//     if (ret === false) {
//         break;
//     }
//     await new Promise(resolve => setTimeout(resolve, (1000)));// aguardar 2 segundos
// }
