let messageCount = 0;

const clearConsole = console.log;

console.log = async function () {

    clearConsole.apply(console, arguments);
    messageCount++;
    if (messageCount >= 300) {
        console.clear();
        messageCount = 0;
        console.log('CONSOLE LIMPO!')
    }

};

export { clearConsole, messageCount };