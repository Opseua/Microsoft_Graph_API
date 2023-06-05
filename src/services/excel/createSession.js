const fs = await import('fs');
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);
const { api } = await import('../../resources/api.js');
const { refreshToken } = await import('../refreshToken.js');

async function createSession() {
  let ret = { 'ret': false };
  const retRefreshToken = await refreshToken();
  if (!retRefreshToken.ret) {
    ret['msg'] = `ERRO AO ATUALIZAR TOKEN`;
  } else {
    if (Date.now() < (config.expireInSession)) {
      ret['msg'] = `SESSAO VALIDA`;
      ret['res'] = { 'token': retRefreshToken.res.token, 'session': config.session };
      ret['ret'] = true;
    } else {
      const fileId = config.fileId;
      const token = retRefreshToken.res.token;
      const corpo = { 'persistChanges': true };
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
      if ('persistChanges' in res) {
        config.session = res.id;
        config.expireInSession = Date.now() + (5 * 60000); // + (5 minutos de validade TESTE)
        fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
        ret['msg'] = `OK CREATE SESSION`;
        ret['res'] = { 'token': retRefreshToken.res.token, 'session': res.id };
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
export { createSession }
