// const infFileWrite = {
//     'file': 'PASTAS 1/PASTA 2/arquivo.txt',
//     'rewrite': true, // 'true' adiciona no MESMO arquivo, 'false' cria outro em branco
//     'text': 'LINHA 1\nLINHA 2\nLINHA 3\n'
// };

// fileWrite(infFileWrite);


import { nodeOrBrowser } from './nodeOrBrowser.js';
import { fileRead } from './fileRead.js';

async function fileWrite(inf) {
    const ret = { ret: false };

    if (inf.file == undefined || inf.file == '') {
        ret['msg'] = `INFORMAR O "file"`;
    } else if (typeof inf.rewrite !== 'boolean') {
        ret['msg'] = `INFORMAR O "rewrite" TRUE ou FALSE`;
    } else if (inf.text == undefined || inf.text == '') {
        ret['msg'] = `INFORMAR O "text"`;
    } else {
        try {
            const resNodeOrBrowser = await nodeOrBrowser()
            if (resNodeOrBrowser.res == 'node') {
                // NODEJS
                const fs = await import('fs');
                const path = await import('path');
                async function createDirectoriesRecursive(directoryPath) {
                    const normalizedPath = path.normalize(directoryPath);
                    const directories = normalizedPath.split(path.sep);
                    let currentDirectory = '';
                    for (let directory of directories) {
                        currentDirectory += directory + path.sep;
                        if (!fs.existsSync(currentDirectory)) { await fs.promises.mkdir(currentDirectory); }
                    }; return true;
                }
                const folderPath = path.dirname(infFileWrite.file);
                await createDirectoriesRecursive(folderPath);
                await fs.promises.writeFile(infFileWrite.file, infFileWrite.text, { flag: infFileWrite.rewrite ? 'a' : 'w' });
            } else {
                // CHROME
                let textOk = inf.text;
                if (inf.rewrite) {
                    const retFileRead = await fileRead(`D:/Downloads/Google Chrome/${inf.file}`)
                    if (retFileRead.ret) { textOk = `${retFileRead.res}${textOk}` }
                }
                const blob = new Blob([textOk], { type: 'text/plain' });
                const downloadOptions = {
                    url: URL.createObjectURL(blob),
                    filename: inf.file,
                    saveAs: false, // PERGUNTAR AO USUARIO ONDE SALVAR
                    conflictAction: 'overwrite' // overwrite (SUBSTITUIR) OU uniquify (REESCREVER→ ADICIONANDO (1), (2), (3)... NO FINAL)
                };
                chrome.downloads.download(downloadOptions, async function (downloadId) {
                    const deleteListDownload = true;
                    await new Promise(resolve => setTimeout(resolve, 2000));
                    if (deleteListDownload) {
                        // EXCLUIR O DOWNLOAD DA LISTA DO CHROME
                        chrome.downloads.erase({ id: downloadId }, function () {
                            //console.log('Download excluído com sucesso');
                        });
                    } else {
                        //console.log('Download não foi excluído');
                    }
                });
            }
            
            ret['ret'] = true;
            ret['msg'] = `FILE WRITE: OK`;
        } catch (e) {
            ret['msg'] = `FILE WRITE: ERRO | ${e}`;
        }
    }

    if (!ret.ret) { console.log(ret.msg) }
    return ret;
}

export { fileWrite };





