const imp1 = () => import('fs').then(module => module.default);
const fs = await imp1();
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

const api = async (i) => (await import('../../resources/api.js')).default(i);

const refreshToken = async (i) => (await import('../refreshToken.js')).default(i);

async function getSheetInf(inf) {
    let ret = { 'ret': false };
    const retRefreshToken = await refreshToken();
    if (!retRefreshToken) {
        ret['msg'] = `ERRO AO ATUALIZAR TOKEN`;
    } else {
        const fileId = config.fileId;
        const token = config.token;
        let sheetTabName = inf.sheetTabName;
        let sheetTabId = '';
        const requisicao = {
            url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets`,
            method: 'GET',
            headers: {
                'authorization': `Bearer ${token}`
            }
        };
        const retApi = await api(requisicao);
        const res = JSON.parse(retApi.res);
        if ('value' in res) {
            if (sheetTabName in config) {
                sheetTabId = config[sheetTabName];
                ret['sheetTabId'] = sheetTabId;
                ret['msg'] = `ABA NOME: ${sheetTabName} | ABA ID: ${sheetTabId}`;
                ret['ret'] = true;
            } else {
                const matchingObject = res.value.find(function (obj) {
                    return obj.name === sheetTabName;
                });
                sheetTabId = matchingObject ? matchingObject.id : null;
                if (sheetTabId == null) {
                    ret['msg'] = `ABA "${sheetTabName}" NAO ENCONTRADA`;
                } else {
                    const novaChave = sheetTabName;
                    const valorNovaChave = sheetTabId;
                    config[novaChave] = valorNovaChave;
                    fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
                    await new Promise(resolve => setTimeout(resolve, (2000)));// aguardar 2 segundos
                    ret['sheetTabId'] = sheetTabId;
                    ret['msg'] = `ABA NOME: ${sheetTabName} | ABA ID: ${sheetTabId}`;
                    ret['ret'] = true;

                }
            }
        } else {
            ret['msg'] = `ERRO AO BUSCAR INFORMACOES DA PLANILHA`;
        }
    }

    console.log(ret.msg);
    return ret
}
export default getSheetInf

