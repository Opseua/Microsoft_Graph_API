async function api(infOk) {
  const inf = {
    url: infOk.url,
    method: infOk.method,
    headers: infOk.headers,
    body: infOk.body
  };

  const ret = { ret: false };

  if (typeof UrlFetchApp !== 'undefined') { // ################ Google App Script
    try {
      var req = UrlFetchApp.fetch(inf.url, {
        'method': inf.method,
        'payload': inf.method === 'POST' || inf.method === 'PATCH' ? typeof inf.body === 'object' ? JSON.stringify(inf.body) : inf.body : null,
        'headers': inf.headers,
        redirect: 'follow',
        keepalive: true,
        muteHttpExceptions: true,
        validateHttpsCertificates: true,
      });
      //console.log('API: OK');
      ret['ret'] = true;
      ret['res'] = req.getContentText();
    } catch (error) {
      console.log('API: ERRO');
      //return error.message;
      ret['res'] = error.message;
    }
  } else { // ######################################### NodeJS ou Chrome
    try {
      var req = await fetch(inf.url, {
        method: inf.method,
        body: inf.method === 'POST' || inf.method === 'PATCH' ? typeof inf.body === 'object' ? JSON.stringify(inf.body) : inf.body : null,
        headers: inf.headers,
        redirect: 'follow',
        keepalive: true
      });
      //console.log('API: OK');
      ret['ret'] = true;
      ret['res'] = await req.text();
    } catch (error) {
      console.log('API: ERRO');
      //return error;
      ret['res'] = error;
    }
  }

  return ret
}
export { api }





/* async function teste() {
  const infApi = {
    url: 'https://ntfy.sh/OPSEUA',
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: String.raw`ESSA \ É / " A ' INFORMACAO`
  };
  const retApi = await api(infApi);
  console.log(retApi.res)
}
teste() */



/* async function teste() {
  // POST → x-www-form-urlencoded
  const formData = new URLSearchParams();
  formData.append('grant_type', 'client_credentials');
  formData.append('client_id', 'c683fd93-XXXXXXXXXXXXX');
  formData.append('client_secret', '0008Q~XXXXXXXXXX');
  formData.append('resource', 'https://graph.microsoft.com');
  const infApi = {
    url: 'https://login.microsoft.com/c5a6c78e-7c99-4631-bb7f-27660b938469/oauth2/token',
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: formData.toString()
  };
  const retApi = await api(infApi);
  console.log(retApi.res)
}
teste() */