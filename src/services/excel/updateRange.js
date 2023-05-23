import { config } from 'dotenv';
config();

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
const fileId = process.env.FILE_ID;
const sheetTabId = process.env.SHEET_TAB_ID;
const oAuth = process.env.TOKEN;

// INFORMACOES EM JSON (fixo)
async function updateRange(inf) {

    let session = '';
    if (inf === undefined) {
        session = process.env.SESSION;
    } else {
        session = inf
    }

    const requisicao = {
        url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets/${sheetTabId}/range(address='A2')`,
        method: 'GET',
        headers: {
            'authorization': `Bearer ${oAuth}`,
            'workbook-session-id': `${session}`
        }
    };
    const re = await api(requisicao);
    const res = JSON.parse(re);
    //console.log("\n\n");
    //console.log(re);
    //console.log("\n\n");
    return res;
}

async function run() {
    const resultado = await updateRange()
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