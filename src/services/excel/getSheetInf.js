import { Console } from 'console';
import fs from 'fs';
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

let fun;
async function api(inf) {
    const module = await import('../../resources/api.js');
    fun = module.default;
    return await fun(inf);
}

let funRefreshToken;
async function refreshToken(inf) {
    const module = await import('../../services/refreshToken.js');
    funRefreshToken = module.default;
    return await funRefreshToken(inf);
}

async function getSheetInf() {
    let msg = '';
    let ret = false;
    const retRefreshToken = await refreshToken();
    if (!retRefreshToken) {
        msg = 'ERRO AO ATUALIZAR TOKEN';
    } else {
        const fileId = config.fileId;
        const token = config.token;
        const requisicao = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets`,
            method: 'GET',
            headers: {
                'authorization': `Bearer ${token}`
            }
        };
        let res = await api(requisicao);
        res = JSON.parse(res);
        if ("value" in res) {
            config.sheetTabId = res.value[1].id;
            config.sheetTabName = res.value[1].name;
            fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
            await new Promise(resolve => setTimeout(resolve, (2000)));// aguardar 2 segundos
            msg = `ABA NOME: ${res.value[1].name} | ABA ID: ${res.value[1].id}`;
            ret = true;
        } else {
            msg = 'ERRO AO BUSCAR INFORMACOES DA PLANILHA';
        }
    }

    console.log(msg);
    return ret
}
export default getSheetInf
