const imp1 = () => import('fs').then(module => module.default);
const fs = await imp1();
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

const api = async (i) => (await import('../../resources/api.js')).default(i);

const refreshToken = async (i) => (await import('../refreshToken.js')).default(i);

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

