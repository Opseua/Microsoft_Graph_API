const fs = await import('fs');
const { fileInf } = await import('../../../../Chrome_Extension/src/resources/fileInf.js');
const retfileInf = await fileInf(new URL(import.meta.url).pathname);
const configPath = `${retfileInf.res.pathProject1}\\src\\config.json`
const configFile = fs.readFileSync(configPath);
const config = JSON.parse(configFile);
const { api } = await import('../../../../Chrome_Extension/src/resources/api.js');
const { refreshToken } = await import('../refreshToken.js');

async function createSession() {
  let ret = { 'ret': false };

  try {
    const retRefreshToken = await refreshToken();
    if (!retRefreshToken.ret) { return ret }

    if (Date.now() < (config.expireInSession)) {
      ret['ret'] = true;
      ret['msg'] = `SESSAO VALIDA`;
      ret['res'] = { 'token': retRefreshToken.res.token, 'session': config.session };
    } else {
      const fileId = config.fileId;
      const token = retRefreshToken.res.token;
      const corpo = { 'persistChanges': true };
      const infApi = {
        url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/createSession`,
        method: 'POST',
        headers: {
          'Content-Type': 'Application/Json',
          'authorization': `Bearer ${token}`
        },
        body: corpo
      };
      const retApi = await api(infApi);
      if (!retApi.ret) { return ret }

      const res = JSON.parse(retApi.res.body);
      if ('persistChanges' in res) {
        config.session = res.id;
        config.expireInSession = Date.now() + (5 * 60000); // + (5 minutos de validade TESTE)
        fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
        ret['ret'] = true;
        ret['msg'] = `OK CREATE SESSION`;
        ret['res'] = { 'token': retRefreshToken.res.token, 'session': res.id };
      }
      else {
        ret['msg'] = `${res.error.code}`;
      }
    }

  } catch (e) {
    ret['msg'] = regexE({ 'e': e }).res;
  }

  if (!ret.ret) { console.log(ret.msg) }
  return ret
}

export { createSession }
