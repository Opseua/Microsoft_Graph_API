const imp1 = () => import('fs').then(module => module.default);
const fs = await imp1();

function writeFile(inf) {

    if (inf.path == undefined || inf.path == '') {
        console.log('INFORMAR O "path"');
        return
    } else if (typeof inf.replace !== 'boolean') {
        console.log('INFORMAR O "replace" TRUE ou FALSE');
        return
    } else if (inf.text == undefined || inf.text == '') {
        console.log('INFORMAR O "text"');
        return
    }

    fs.writeFile(inf.path, inf.text, { flag: inf.replace ? 'w' : 'a' }, function (err) {
        if (err) {
            console.error('ARQUIVO: ERRO | ', err);
        } else {
            console.log('ARQUIVO: OK');
        }
    });

}

export default writeFile


import path from 'path';

const projectRoot = path.resolve();

console.log(projectRoot);