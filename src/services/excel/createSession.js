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

async function createSession() {
  let ret = { 'ret': false };
  const retRefreshToken = await refreshToken();
  if (!retRefreshToken) {
    ret['msg'] = `ERRO AO ATUALIZAR TOKEN`;
  } else {
    if (Date.now() < (config.expireInSession - 3000)) {
      ret['msg'] = `SESSAO VALIDA`;
      ret['ret'] = true;
    } else {
      const fileId = config.fileId;
      const token = config.token;
      const corpo = { "persistChanges": true };
      const requisicao = {
        url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/createSession`,
        method: 'POST',
        headers: {
          'Content-Type': 'Application/Json',
          'authorization': `Bearer ${token}`
        },
        body: corpo
      };
      const retApi = await api(requisicao);
      const res = JSON.parse(retApi.res);
      if ("persistChanges" in res) {
        config.session = res.id;
        config.expireInSession = Date.now() + (10 * 60000); // + X minutos
        fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
        await new Promise(resolve => setTimeout(resolve, (2000)));// aguardar 2 segundos
        ret['msg'] = `OK CREATE SESSION`;
        ret['ret'] = true;
      }
      else {
        ret['msg'] = `${res.error.code}`;
      }
    }
  }

  console.log(ret.msg);
  return ret
}
export default createSession
