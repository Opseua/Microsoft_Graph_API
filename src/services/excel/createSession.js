import fs from 'fs';
const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

let fun;
async function api(inf) {
  const module = await import('../../resources/api.js');
  fun = module.default;
  return await fun(inf);
}

const fileId = config.fileId
const token = config.token

async function createSession() {
  const corpo = { "persistChanges": true }
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
  let msg = ''
  if ("persistChanges" in res) {
    config.session = res.id;
    fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
    msg = 'OK CREATE SESSION'
  }
  else {
    msg = res.error.code;
  }
  console.log(msg)
}
export default createSession

createSession()