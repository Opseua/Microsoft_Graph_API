import { config } from 'dotenv';
config();

let fun;
async function api(inf) {
  const module = await import('../../resources/api.js');
  fun = module.default;
  return await fun(inf);
}

const fileId = process.env.FILE_ID;
const oAuth = process.env.TOKEN;

async function createSession() {
  const corpo = { "persistChanges": true }
  const requisicao = {
    url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/createSession`,
    method: 'POST',
    headers: {
      'Content-Type': 'Application/Json',
      'authorization': `Bearer ${oAuth}`
    },
    body: corpo
  };
  let res = await api(requisicao);
  res = JSON.parse(res);
  const session = res.id;
  //console.log("\n\n");
  //console.log(session);
  //console.log("\n\n");
  return session;
}
createSession()

export default createSession