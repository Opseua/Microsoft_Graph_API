const fs = await import('fs');
const { fileInf } = await import('../../../../Chrome_Extension/src/resources/fileInf.js');
const retfileInf = await fileInf(new URL(import.meta.url).pathname);
const configPath = `${retfileInf.res.pathProject1}\\src\\config.json`
const configFile = fs.readFileSync(configPath);
const config = JSON.parse(configFile);
const { api } = await import('../../../../Chrome_Extension/src/resources/api.js');
const { refreshToken } = await import('../refreshToken.js');

async function getSheetInf(inf) {
    let ret = { 'ret': false };
    try {
        const retRefreshToken = await refreshToken();
        if (!retRefreshToken.ret) { return ret }

        let sheetTabName = inf.sheetTabName
        if (sheetTabName in config) {
            ret['res'] = { 'token': retRefreshToken.res.token, 'sheetTabId': config[sheetTabName] };
            ret['msg'] = `ABA NOME: ${sheetTabName} | ABA ID: ${config[sheetTabName]}`;
            ret['ret'] = true;
        } else {
            const fileId = config.fileId;
            const token = retRefreshToken.res.token;
            let sheetTabId
            const infApi = {
                url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets`,
                method: 'GET',
                headers: {
                    'authorization': `Bearer ${token}`
                }
            };
            const retApi = await api(infApi);
            if (!retApi.ret) { return ret }

            const res = JSON.parse(retApi.res.body);
            if ('value' in res) {
                const matchingObject = res.value.find(function (obj) {
                    return obj.name === sheetTabName;
                });
                sheetTabId = matchingObject ? matchingObject.id : null;
                if (sheetTabId == null) {
                    ret['msg'] = `ABA "${sheetTabName}" NAO ENCONTRADA`;
                } else {
                    config[sheetTabName] = sheetTabId;
                    fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
                    ret['res'] = { 'token': retRefreshToken.res.token, 'sheetTabId': sheetTabId };
                    ret['msg'] = `ABA NOME: ${sheetTabName} | ABA ID: ${sheetTabId}`;
                    ret['ret'] = true;
                }
            } else {
                ret['msg'] = `ERRO AO BUSCAR INFORMACOES DA PLANILHA`;
            }
        }
    } catch (e) {
        ret['msg'] = regexE({ 'e': e }).res;
    }

    if (!ret.ret) { console.log(ret.msg) }
    return ret
}

export { getSheetInf }

