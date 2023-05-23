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

const fileId = config.fileId
const sheetTabName = config.sheetTabName
const sheetRange = config.sheetRange
const token = config.token
const session = config.session

// INFORMACOES EM JSON (fixo)
async function updateRange(inf) {

    let session = '';
    if (inf === undefined) {
        session = config.session
    } else {
        session = inf;
    }

    const corpo = { "values": [["Test"]] };
    const requisicao = {
        url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheetTabName}')/range(address='${sheetRange}')`,
        method: 'PATCH',
        headers: {
            'Content-Type': 'Application/Json',
            'authorization': `Bearer ${token}`,
            'workbook-session-id': `${session}`
        },
        body: corpo
    };
    const re = await api(requisicao);
    console.log("aaaaaaa")
    console.log(re)
    //const res = JSON.parse(re);
    //console.log(res)
    //return res;
    return re;
}

async function run() {
    const resultado = await updateRange()

    console.log(resultado)
    return

    if (resultado.error.code == 'undefined') {
        console.log(resultado);
    }
    else if (resultado.error.code == 'InvalidAuthenticationToken') {
        console.log('TOKEN INVALIDO');
    }
    else if (resultado.error.code == 'InvalidSession') {
        console.log('SESSAO INVALIDA');
        const session = await createSession();
        const resultado = await updateRange(session);
        console.log(resultado.values[0]);
    } else {
        console.log("OUTRO ERRO");
    }

}
run()