import fs from 'fs';
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

let funApi;
async function api(inf) {
    const module = await import('../../resources/api.js');
    funApi = module.default;
    return await funApi(inf);
}

let funRefreshToken;
async function refreshToken(inf) {
    const module = await import('../../services/refreshToken.js');
    funRefreshToken = module.default;
    return await funRefreshToken(inf);
}

async function listAllFilesByType() {
    let ret = {
        'ret': false
    };
    const retRefreshToken = await refreshToken();
    if (!retRefreshToken) {
        ret['msg'] = `ERRO AO ATUALIZAR TOKEN`;
    } else {
        const query = config.fileName;
        const token = config.token;
        const requisicao = {
            url: `https://graph.microsoft.com/v1.0/me/drive/root/search(q=\'${query}\')`,
            method: 'GET',
            headers: {
                'authorization': `Bearer ${token}`
            }
        };
        let res = await api(requisicao);
        res = JSON.parse(res);
        if ("value" in res) {
            if (res.value.length == 0) {
                ret['msg'] = `NENHUM ARQUIVO ENCONTRADO`;
            } else {
                config.fileId = res.value[0].id;
                fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
                await new Promise(resolve => setTimeout(resolve, (2000)));// aguardar 2 segundos
                ret['msg'] = `ARQUIVO NOME: ${res.value[0].name} | ARQUIVO ID: ${res.value[0].id}`;
                ret['ret'] = true;
            }
        }
        else {
            ret['msg'] = `${res.error.code}`;
        }
    }

    console.log(ret.msg);
    return ret
}
export default listAllFilesByType
listAllFilesByType()