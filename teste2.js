let teste1;
async function iniciarExtensao(inf) {
    const module = await import('./teste1.js');
    teste1 = module.default;
    teste1(inf);
}
iniciarExtensao("aaaaa");









