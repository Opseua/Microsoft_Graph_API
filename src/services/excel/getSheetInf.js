const fs = await import('fs');
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);
const { api } = await import('../../resources/api.js');
const { refreshToken } = await import('../refreshToken.js');

async function getSheetInf(inf) {
    let ret = { 'ret': false };
    const retRefreshToken = await refreshToken();
    if (!retRefreshToken.ret) {
        ret['msg'] = `ERRO AO ATUALIZAR TOKEN`;
    } else {
        let sheetTabName = inf.sheetTabName
        if (sheetTabName in config) {
            ret['res'] = { 'token': retRefreshToken.res.token, 'sheetTabId': config[sheetTabName] };
            ret['msg'] = `ABA NOME: ${sheetTabName} | ABA ID: ${config[sheetTabName]}`;
            ret['ret'] = true;
        } else {
            const fileId = config.fileId;
            const token = retRefreshToken.res.token;
            let sheetTabId
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
                const matchingObject = res.value.find(function (obj) {
                    return obj.name === sheetTabName;
                });
                sheetTabId = matchingObject ? matchingObject.id : null;
                if (sheetTabId == null) {
                    ret['msg'] = `ABA "${sheetTabName}" NAO ENCONTRADA`;
                } else {
                    config[sheetTabName] = sheetTabId;
                    fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
                    ret['res'] = { 'token': retRefreshToken.res.token, 'sheetTabId': sheetTabId };
                    ret['msg'] = `ABA NOME: ${sheetTabName} | ABA ID: ${sheetTabId}`;
                    ret['ret'] = true;
                }
            } else {
                ret['msg'] = `ERRO AO BUSCAR INFORMACOES DA PLANILHA`;
            }
        }
    }

    console.log(ret.msg);
    return ret
}
export { getSheetInf }

