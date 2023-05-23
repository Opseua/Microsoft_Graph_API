import fs from 'fs';
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

let fun;
async function api(inf) {
    const module = await import('../resources/api.js');
    fun = module.default;
    return await fun(inf);
}

const clientId = config.clientId
const refresh = config.refresh

async function refreshToken() {
    const formData = new URLSearchParams();
    formData.append('client_id', clientId);
    formData.append('refresh_token', refresh);
    formData.append('grant_type', 'REFRESH_TOKEN');
    const corpo = formData.toString()
    const requisicao = {
        url: `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: corpo
    };
    let res = await api(requisicao);
    res = JSON.parse(res);
    let msg = ''
    if ("access_token" in res) {
        config.token = res.access_token;
        config.refresh = res.refresh_token;
        fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
        msg = 'OK REFRESH TOKEN'
    }
    else {
        msg = res.error;
    }
    console.log(msg)
}
export default refreshToken

refreshToken()

