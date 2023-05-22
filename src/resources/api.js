async function api(inf_ok) {

  const inf = {
    url: inf_ok.url,
    method: inf_ok.method,
    headers: inf_ok.headers,
    body: inf_ok.body
  };

  if (typeof fetch == "undefined") {
    // Google App Script
    try {
      var req = UrlFetchApp.fetch(inf.url, {
        'method': inf.method,
        'payload': inf.method === "POST" ? typeof inf.body === 'object' ? JSON.stringify(inf.body) : inf.body : null,
        'headers': inf.headers,
        redirect: 'follow',
        keepalive: true,
        muteHttpExceptions: true,
        validateHttpsCertificates: true,
      });
      console.log('API: OK');
      return req.getContentText();
    } catch (error) {
      console.log('API: ERRO');
      return error.message;
    }
  } else {
    // JavaScript
    try {
      var req = await fetch(inf.url, {
        method: inf.method,
        body: inf.method === "POST" ? typeof inf.body === 'object' ? JSON.stringify(inf.body) : inf.body : null,
        headers: inf.headers,
        redirect: 'follow',
        keepalive: true
      });
      console.log('API: OK');
      return await req.text();
    } catch (error) {
      console.log('API: ERRO');
      return error;
    }
  }

}

export default api





/* async function teste() {
  const corpo = String.raw`ESSA \ É / " A ' INFORMACAO`;
  //const corpo = { teste: 'OLA TUDO BEM' };
  const requisicao = {
    url: 'https://ntfy.sh/OPSEUA',
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: valor
  };
  const re = await api(requisicao);
  console.log(re)
}
teste() */



/* async function teste() {
  // POST → x-www-form-urlencoded
  const formData = new URLSearchParams();
  formData.append('grant_type', 'client_credentials');
  formData.append('client_id', 'c683fd93-XXXXXXXXXXXXX');
  formData.append('client_secret', '0008Q~XXXXXXXXXX');
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
teste() */