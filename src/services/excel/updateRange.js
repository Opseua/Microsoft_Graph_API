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

let celula = 'TESTE!B';
let numero = 10;

// INFORMACOES EM JSON (fixo)
async function updateRange(inf) {

    let session = '';
    if (inf === undefined) {
        session = config.session
    } else {
        session = inf;
    }

    numero += 1;

    const corpo = { "values": [[`${celula}${numero}`]] };
    const requisicao = {
        url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheetTabName}')/range(address='${celula}${numero}')`,
        method: 'PATCH',
        headers: {
            'Content-Type': 'Application/Json',
            'authorization': `Bearer ${token}`,
            'workbook-session-id': `${session}`
        },
        body: corpo
    };
    const re = await api(requisicao);
    const res = JSON.parse(re);
    return res;
}

async function run() {
    let msg = '';
    let resultado = await updateRange()
    if ("values" in resultado) {
        msg = resultado.values[0];
    }
    else if (resultado.error.code == 'InvalidSession') {
        const session = await createSession();
        resultado = await updateRange(session);
        if ("values" in resultado) {
            msg = resultado.values[0];
        } else {
            msg = resultado.error.message;
        }
    } else {
        msg = resultado.error.message;
    }
    console.log(msg)
}



function executarLoopComAtraso() {
    for (let i = 0; i < 20; i++) {
        setTimeout(function () {
            run()
        }, i * 3000); // Multiplica o índice da iteração pelo tempo de atraso (5000 ms = 5 segundos)
    }
}

// Chamar a função para iniciar o loop com atraso
executarLoopComAtraso();
