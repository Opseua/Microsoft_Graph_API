import { config } from 'dotenv';
config();

let fun;
async function api(inf) {
    const module = await import('../resources/api.js');
    fun = module.default;
    return await fun(inf);
}

const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const clientDirectory = process.env.CLIENT_DIRECTORY;

async function teste() {
    // POST → x-www-form-urlencoded
    const formData = new URLSearchParams();
    formData.append('grant_type', 'client_credentials');
    formData.append('client_id', clientId);
    formData.append('client_secret', clientSecret);
    formData.append('resource', 'https://graph.microsoft.com');
    const corpo = formData.toString()
    const requisicao = {
        url: `https://login.microsoft.com/${clientDirectory}/oauth2/token`,
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: corpo
    };
    const re = await api(requisicao);
    const res = JSON.parse(re);
    process.env.OAUTH_TOKEN = res.access_token;
    console.log("\n\n");
    console.log(process.env.OAUTH_TOKEN);
    console.log("\n\n");
}
teste()

