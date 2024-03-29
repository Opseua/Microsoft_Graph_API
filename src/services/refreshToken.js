await import('../../../Chrome_Extension/src/clearConsole.js');
const fs = await import('fs');
const { fileInf } = await import('../../../Chrome_Extension/src/resources/fileInf.js');
const retfileInf = await fileInf(new URL(import.meta.url).pathname);
const configPath = `${retfileInf.res.pathProject1}\\src\\config.json`
const configFile = fs.readFileSync(configPath);
const config = JSON.parse(configFile);
const { api } = await import('../../../Chrome_Extension/src/resources/api.js');

async function refreshToken() {
    let ret = { 'ret': false };
    try {
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
            const infApi = {
                url: `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: formData.toString()
            };
            const retApi = await api(infApi);
            if (!retApi.ret) { return ret }

            const res = JSON.parse(retApi.res.body);
            if ('access_token' in res) {
                config.token = res.access_token;
                config.refresh = res.refresh_token;
                config.expireInRefresh = Date.now() + ((res.expires_in - 60) * 1000); // + (59 minutos de validade)
                fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
                ret['msg'] = `OK REFRESH TOKEN`;
                ret['res'] = { 'token': res.access_token };
                ret['ret'] = true;
            }
            else {
                ret['msg'] = `${res.error}`;
                ret['ret'] = false;
            }
        }
    } catch (e) {
        ret['msg'] = regexE({ 'e': e }).res;
    }

    if (!ret.ret) { console.log(ret.msg) }
    return ret
}
export { refreshToken }
