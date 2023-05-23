import { config } from 'dotenv';
config();

let fun;
async function api(inf) {
    const module = await import('../../resources/api.js');
    fun = module.default;
    return await fun(inf);
}
const oAuth = process.env.OAUTH_TOKEN;
const fileId = process.env.FILE_ID;

async function getSheetInf() {
    const requisicao = {
        url: `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets`,
        method: 'GET',
        headers: {
            'authorization': `Bearer ${oAuth}`
        }
    };
    const re = await api(requisicao);
    const res = JSON.parse(re);
    process.env.TABSHEETID = res.value[0].id;
    process.env.TABSHEETNAME = res.value[0].name;
    console.log("\n\n");
    console.log(res.value[0].name);
    console.log(res.value[0].id);
    console.log("\n\n");
}
getSheetInf()