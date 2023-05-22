import { config } from 'dotenv';
config();

let fun;
async function api(inf) {
    const module = await import('../../resources/api.js');
    fun = module.default;
    return await fun(inf);
}
const oAuth = process.env.OAUTH_TOKEN;
const query = 'NOME_AQUI.xlsx';

async function listFiles() {
    const requisicao = {
        url: `https://graph.microsoft.com/v1.0/me/drive/root/search(q=\'${query}\')`,
        method: 'GET',
        headers: {
            'authorization': `Bearer ${oAuth}`
        }
    };
    const re = await api(requisicao);
    const res = JSON.parse(re);
    process.env.FILE_ID = res.value[0].id;
    console.log("\n\n");
    console.log(res.value[0].id);
    console.log("\n\n");
}
listFiles()