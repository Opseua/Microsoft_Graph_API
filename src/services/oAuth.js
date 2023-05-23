import { config } from 'dotenv';
config();

let fun;
async function api(inf) {
    const module = await import('../resources/api.js');
    fun = module.default;
    return await fun(inf);
}

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;


async function oAuth() {

    //console.log(clientDirectory)
    //return

    const formData = new URLSearchParams();
    formData.append('grant_type', 'client_credentials');
    formData.append('client_id', clientId);
    formData.append('client_secret', clientSecret);
    formData.append('resource', 'https://graph.microsoft.com');
    const corpo = formData.toString()
    const requisicao = {
        url: `https://login.microsoft.com/${tenantId}/oauth2/token`,
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: corpo
    };
    let res = await api(requisicao);
    res = JSON.parse(res);
    const token = res.access_token;
    //process.env.OAUTH_TOKEN = token;
    console.log("\n\n");
    console.log(token);
    console.log("\n\n");
}
oAuth()

