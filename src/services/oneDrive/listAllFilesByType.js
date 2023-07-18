const fs = await import('fs');
const { fileInf } = await import('../../../../Chrome_Extension/src/resources/fileInf.js');
const retfileInf = await fileInf(new URL(import.meta.url).pathname);
const configPath = `${retfileInf.res.pathProject1}\\src\\config.json`
const configFile = fs.readFileSync(configPath);
const config = JSON.parse(configFile);
const { api } = await import('../../../../Chrome_Extension/src/resources/api.js');
const { refreshToken } = await import('../refreshToken.js');

async function listAllFilesByType() {
    let ret = { 'ret': false };
    try {
        const retRefreshToken = await refreshToken();
        if (!retRefreshToken.ret) { return ret }

        const query = config.fileName;
        const token = config.token;
        const infApi = {
            url: `https://graph.microsoft.com/v1.0/me/drive/root/search(q=\'${query}\')`,
            method: 'GET',
            headers: {
                'authorization': `Bearer ${token}`
            }
        };
        const retApi = await api(infApi);
        if (!retApi.ret) { return ret }

        const res = JSON.parse(retApi.res.body);
        if ('value' in res) {
            if (res.value.length == 0) {
                ret['msg'] = `NENHUM ARQUIVO ENCONTRADO`;
            } else {
                config.fileId = res.value[0].id;
                fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
                await new Promise(resolve => setTimeout(resolve, (2000)));// aguardar 2 segundos
                ret['msg'] = `ARQUIVO NOME: ${res.value[0].name} | ARQUIVO ID: ${res.value[0].id}`;
                ret['ret'] = true;
            }
        }
        else {
            ret['msg'] = `${res.error.code}`;
        }

    } catch (e) {
        ret['msg'] = `LIST ALL FILES BY TYPE: ERRO | ${e}`;
    }

    if (!ret.ret) { console.log(ret.msg) }
    return ret
}

export { listAllFilesByType }
