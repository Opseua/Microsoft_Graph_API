import fs from 'fs';
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

let funApi;
async function api(inf) {
    const module = await import('../../resources/api.js');
    funApi = module.default;
    return await funApi(inf);
}
let funGetRange;
async function getRange(inf) {
    const module = await import('./getRange.js');
    funGetRange = module.default;
    return await funGetRange(inf);
}

let fileId;
let retGetRange
let sheetTabName
let sheetCol
let sheetLin
let token
let session
let run = 0

async function updateRange(inf) {
    let ret = { 'ret': false };
    //console.log(`NUMERO ${sheetLin}`)

    if (run == 0) {
        console.log('COMECOU')
        retGetRange = await getRange({ sheetTabName: inf.sheetTabName });
        if (!retGetRange.ret) {
            run = -1;
            console.log('GET RANGE: ERRO');
        } else {
            run = 1;
            console.log('RODAR');
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
        //console.log(JSON.stringify(requisicao.body))
        //return
        if (!("values" in res)) {
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



for (let i = 0; i < 1; i++) {
    const ret = await updateRange({ sheetTabName: 'HAUPC', send: 'OLÁ' });
    if (ret === false) {
        break;
    }
    await new Promise(resolve => setTimeout(resolve, (1000)));// aguardar 2 segundos
}
