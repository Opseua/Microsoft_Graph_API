import fs from 'fs';

const configFile = fs.readFileSync('config.json');
const config = JSON.parse(configFile);

// Modifica os valores das configurações
config.clientId = 'NOVO clientId';
config.tenantId = 'NOVO tenantId';
config.clientSecret = "NOVO clientSecret";
config.sheetTabName = "NOVO sheetTabName";
config.sheetTabId = "NOVO sheetTabId";
config.token = "NOVO token";
config.session = "NOVO session";
config.refresh = "NOVO refresh";

// Salva as alterações no arquivo de configuração
fs.writeFileSync('config.json', JSON.stringify(config, null, 2));

// Utiliza os valores das configurações
//console.log(config.apiKey);
//console.log(config.baseUrl);
//console.log(config.timeout);


