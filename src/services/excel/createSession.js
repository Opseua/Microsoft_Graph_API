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
  let msg = '';
  let ret = false;
  const retRefreshToken = await refreshToken();
  if (!retRefreshToken) {
    let msg = 'ERRO AO ATUALIZAR TOKEN';
  } else {
    if (Date.now() < (config.expireInSession - 3000)) {
      msg = 'SESSAO VALIDA';
      ret = true;
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
      let res = await api(requisicao);
      res = JSON.parse(res);
      if ("persistChanges" in res) {
        config.session = res.id;
        config.expireInSession = Date.now() + 1800000; // + 30 minutos
        fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
        msg = 'OK CREATE SESSION';
        ret = true;
      }
      else {
        msg = res.error.code;
      }
    }
  }

  console.log(msg);
  return ret
}
export default createSession
