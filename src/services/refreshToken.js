
import fs from 'fs';
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

let funApi;
async function api(inf) {
    const module = await import('../resources/api.js');
    funApi = module.default;
    return await funApi(inf);
}

async function refreshToken() {
    let msg = '';
    let ret = false;
    if (Date.now() < (config.expireInRefresh - 3000)) {
        msg = 'TOKEN VALIDO';
        ret = true;
    } else {
        const clientId = config.clientId;
        const refresh = config.refresh;
        const formData = new URLSearchParams();
        formData.append('client_id', clientId);
        formData.append('refresh_token', refresh);
        formData.append('grant_type', 'REFRESH_TOKEN');
        const corpo = formData.toString();
        const requisicao = {
            url: `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: corpo
        };
        let res = await api(requisicao);
        res = JSON.parse(res);
        if ("access_token" in res) {
            config.token = res.access_token;
            config.refresh = res.refresh_token;
            config.expireInRefresh = Date.now() + (res.expires_in * 1000); // + 1 hora
            fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
            msg = 'OK REFRESH TOKEN';
            ret = true;
        }
        else {
            msg = res.error;
            ret = false;
        }
    }

    console.log(msg);
    return ret
}
export default refreshToken



