/* let Api;
import('./src/recursos/Api.js').then(module => {
    Api = module.default;
});
 */

const { api } = require('./api.js');


async function teste() {
    // POST → x-www-form-urlencoded
    const formData = new URLSearchParams();
    formData.append('grant_type', 'client_credentials');
    formData.append('client_id', 'c683fd93-7db0-44a5-98a3-7fe0890b5c90');
    formData.append('client_secret', '0008Q~5O2MA2Q-YeMJQN9de_4pPVzx8tse096c38');
    formData.append('resource', 'https://graph.microsoft.com');
    const corpo = formData.toString()
    const requisicao = {
        url: 'https://login.microsoft.com/c5a6c78e-7c99-4631-bb7f-27660b938469/oauth2/token',
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: corpo
    };
    const re = await api(requisicao);
    console.log(re)
}
teste()