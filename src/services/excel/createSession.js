import { config } from 'dotenv';
config();

let fun;
async function api(inf) {
  const module = await import('../../resources/api.js');
  fun = module.default;
  return await fun(inf);
}

const fileId = process.env.FILE_ID;
const oAuth = process.env.OAUTH_TOKEN;

async function getSession() {
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
  const re = await api(requisicao);
  const res = JSON.parse(re);
  process.env.SESSION = res.id;
  console.log("\n\n");
  console.log(res.id);
  console.log("\n\n");
}
getSession()
