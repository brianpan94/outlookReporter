let CONFIGS = require('./configs.js');
let func = require('./functions.js');
let argv = require('minimist')(process.argv.slice(2));
delete argv._;

let { start, end, subject, fullSubject, to, 
    message, token, pushId,
    // pushRetry, pushRequestOnce, pushTimeout,  // no use here
    workRetry, workRetryTimeout} = argv;

CONFIGS.start = start ? start : CONFIGS.start; 
CONFIGS.end = end ? end : CONFIGS.end;
CONFIGS.workRetry = workRetry ? workRetry : CONFIGS.workRetry; 
CONFIGS.workRetryTimeout = workRetryTimeout ? workRetryTimeout : CONFIGS.workRetryTimeout;

main();

async function main(retry){
    retry = retry || 0;
    console.log(`${taiwanTime()} starting ${ retry++ }st time`);
    console.log('params:', argv);

    func.getTargetEmail(CONFIGS.start, CONFIGS.end, subject, to, fullSubject || 'N')
    // .then(response => {
    //     return func.getEmailContent(response);
    // })
    .then(response => {
        return func.filterContent(response);
    })
    .then(response => {
        // console.log(unicodeToChinese(message), token, response, pushId, "=====/n/n");
        // return;
        return func.pushToTarget(unicodeToChinese(message), token, response, pushId);
        // return func.broadcastToMembers(message, token, response);
    })
    .then(response => {
        console.log(taiwanTime() + ' finish task');
    })
    .catch(error => {
        console.log(taiwanTime() + ' fail: ' + error);
        console.log(`next try will start at: ${ taiwanTime(CONFIGS.workRetryTimeout) }`);
        if(retry < CONFIGS.workRetry){
            setTimeout(() => {
                main(retry);
            }, CONFIGS.workRetryTimeout);
        }
    })
    .finally(() => {
        console.log('============== task end ==============');
    })
}

function unicodeToChinese(text = '') {
    return text.replace(/\\\u[\dA-F]{4}/gi, 
        function (match) {
            return String.fromCharCode(parseInt(match.replace(/\\u/g, ''), 16));
        });
}

function taiwanTime(offset = 0) {
    let timeZoneOffset = new Date().getTimezoneOffset() / 60;
    return new Date(new Date().getTime() + (-8 - timeZoneOffset) * 60*60*1000 + offset).toLocaleString('zh-Hans-CN');
}