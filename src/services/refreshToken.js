await import('./../clearConsole.js');

const imp1 = () => import('fs').then(module => module.default);
const fs = await imp1();
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

//const api = async (i) => (await import('../resources/api.js')).default(i);
const { api } = await import('../resources/api.js');

async function refreshToken() {
    const ret = { 'ret': false };
    if (Date.now() < (config.expireInRefresh)) {
        ret['msg'] = `TOKEN VALIDO`;
        ret['res'] = { 'token': config.token };
        ret['ret'] = true;
    } else {
        const clientId = config.clientId;
        const refresh = config.refresh;
        const formData = new URLSearchParams();
        formData.append('client_id', clientId);
        formData.append('refresh_token', refresh);
        formData.append('grant_type', 'REFRESH_TOKEN');
        const requisicao = {
            url: `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: formData.toString()
        };
        const retApi = await api(requisicao);
        const res = JSON.parse(retApi.res);
        if ('access_token' in res) {
            config.token = res.access_token;
            config.refresh = res.refresh_token;
            config.expireInRefresh = Date.now() + ((res.expires_in - 60) * 1000); // + (59 minutos de validade)
            fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
            ret['msg'] = `OK REFRESH TOKEN`;
            ret['res'] = { 'token': res.access_token };
            ret['ret'] = true;
        }
        else {
            ret['msg'] = `${res.error}`;
            ret['ret'] = false;
        }
    }

    console.log(ret.msg);
    return ret
}
export default refreshToken

