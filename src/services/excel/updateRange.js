const { addListener, globalObject } = await import('../../../../Chrome_Extension/src/resources/globalObject.js');
addListener(monitorGlobalObject);
async function monitorGlobalObject(value) {
    if (value.inf.function == 'updateRange') {
        //console.log('→→→', value.inf.res)
        await updateRange({ 'sheetTabName': 'YVIE', 'send': value.inf.res });
    }
}

// *****************************************************************

const fs = await import('fs');
import { fileInf } from '../../../../Chrome_Extension/src/resources/fileInf.js';
const retfileInf = await fileInf(new URL(import.meta.url).pathname);
const configPath = `${retfileInf.res.pathProject1}\\config.json`
const configFile = fs.readFileSync(configPath);
const config = JSON.parse(configFile);
// const configFile = fs.readFileSync('config.json');
// const config = JSON.parse(configFile);
const { api } = await import('../../resources/api.js');
const { getRange } = await import('./getRange.js');

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
            sheetLin = Number(retGetRange.res.values);
            token = retGetRange.res.token;
            session = retGetRange.res.session;
        }
    }
    if (run == 1) {
        const infApi = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheetTabName}')/range(address='${sheetCol}${sheetLin}')`,
            method: 'PATCH',
            headers: {
                'Content-Type': 'Application/Json',
                'authorization': `Bearer ${token}`,
                'workbook-session-id': `${session}`
            },
            body: { 'values': [[`${sheetLin} | ${inf.send}`]] }
        };
        const retApi = await api(infApi);
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
export { updateRange }



// for (let i = 0; i < 10; i++) {
//     const ret = await updateRange({ 'sheetTabName': 'YVIE', 'send': 'OLÁ' });
//     if (ret === false) {
//         break;
//     }
//     await new Promise(resolve => setTimeout(resolve, (1000)));// aguardar 2 segundos
// }


