const request = require('request-promise');
const HTMLPARSER = require('node-html-parser');
let CONFIGS = require('./configs.js');
let { ClientSecretCredential } = require('@azure/identity');
let { Client } = require('@microsoft/microsoft-graph-client');
let { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let argv = require('minimist')(process.argv.slice(2));
delete argv._;

let { pushRetry, pushRequestOnce, pushTimeout } = argv;

CONFIGS.pushRetry = pushRetry ? pushRetry : CONFIGS.pushRetry;
CONFIGS.pushRequestOnce = pushRequestOnce ? pushRequestOnce : CONFIGS.pushRequestOnce;
CONFIGS.pushTimeout = pushTimeout ? pushTimeout : CONFIGS.pushTimeout;

const { TENANT_ID: tenantId, CLIENT_ID: clientId, CLIENT_SECRET: clientSecret } = CONFIGS.azureClient;
const tokenCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const options = { scopes: ['https://graph.microsoft.com/.default'], options: { tenantId: tenantId } };
const authProvider = new TokenCredentialAuthenticationProvider(tokenCredential, options);
const client = Client.initWithMiddleware({
    debugLogging: true,
    authProvider: authProvider,
});

// 找到指定信件、撈出內文HTML 
// 2023.02.03 更新產出email內容
exports.getTargetEmail = async (start, end, subject, to, fullSubject = 'N') => {

    if (subject && to) return Promise.reject('"subject" and "to" should only exit one');

    let findMailEwsArgs = {
        'attributes': {
            'Traversal': 'Shallow'
        },
        'ItemShape': {
            'BaseShape': 'Default',
            // 'AdditionalProperties': {
            // 	'FieldURI': [
            //         { 'attributes': { 'FieldURI': 'item:DisplayTo' } }, 
            //         { 'attributes': { 'FieldURI': 'item:DisplayCc' } }, 
            //     ]
            // }
        },
        'Restriction': {
            't:And': {}
            // {
            //     // 信件收到時間區間
            // 	't:IsGreaterThanOrEqualTo': {
            // 	  't:FieldURI': { attributes: { FieldURI: 'item:DateTimeReceived' } },
            // 	  't:FieldURIOrConstant': {
            // 	    't:Constant': {attributes: { Value: new Date(start).toISOString() } }
            // 	  }
            // 	},
            // 	't:IsLessThanOrEqualTo': {
            // 	  't:FieldURI': { attributes: { FieldURI: 'item:DateTimeReceived' } },
            // 	  't:FieldURIOrConstant': {
            // 	    't:Constant': { attributes: { Value: new Date(end).toISOString() } }
            // 	  }
            //     },

            // 	// 't:IsEqualTo': {
            // 	// 	't:FieldURI': { attributes: { FieldURI: 'item:DisplayTo' } },
            // 	// 	't:FieldURIOrConstant': {
            // 	// 		't:Constant': { attributes: { Value: 'Chimei專案測試' } }
            // 	// 	}
            //     // },


            //     // 信件標題包含
            // 	't:Contains': {
            // 		'attributes': {
            // 			'ContainmentMode': 'Substring',
            // 			'ContainmentComparison': 'IgnoreCase',
            // 		},
            // 		't:FieldURI': { attributes: { FieldURI: 'item:Subject' } },
            // 		't:Constant': { attributes: { Value: subject } }
            //     },

            //     // 信件標題相符
            // 	// 't:IsEqualTo': {
            // 	//   't:FieldURI': { attributes: { FieldURI: 'item:Subject' } },
            // 	//   't:FieldURIOrConstant': {
            // 	//     't:Constant': { attributes: { Value: 'NBU-CHshow-Daily Report' } }
            // 	//   }
            //     // },

            //     // 收件人包含
            //     // 't:Contains': {
            // 	// 	'attributes': {
            // 	// 		'ContainmentMode': 'Substring',
            // 	// 		'ContainmentComparison': 'IgnoreCase',
            // 	// 	},
            // 	// 	't:FieldURI': { attributes: { FieldURI: 'item:DisplayTo' } },
            // 	// 	't:Constant': { attributes: { Value: to } }
            // 	// },
            // }
        },
        'ParentFolderIds': {
            'DistinguishedFolderId': {
                'attributes': {
                    'Id': 'inbox'
                }
            }
        }
    }
    let filterBody = {};
    if (start) {
        console.log('from:', taiwanUTCString(new Date().toLocaleDateString('zh-Hans-CN') + ' ' + start));
        filterBody.start = taiwanUTCString(new Date().toLocaleDateString('zh-Hans-CN') + ' ' + start);
    }
    if (end) {
        console.log('to:', taiwanUTCString(new Date().toLocaleDateString('zh-Hans-CN') + ' ' + end));
        filterBody.end = taiwanUTCString(new Date().toLocaleDateString('zh-Hans-CN') + ' ' + end);
    }
    if (subject) {
        if (fullSubject == 'Y') {
            // findMailEwsArgs.Restriction['t:And']['t:IsEqualTo'] = {
            //     't:FieldURI': { attributes: { FieldURI: 'item:Subject' } },
            //     't:FieldURIOrConstant': {
            //         't:Constant': { attributes: { Value: subject } }
            //     }
            // }
        } else {
            // findMailEwsArgs.Restriction['t:And']['t:Contains'] = {
            //     'attributes': {
            //         'ContainmentMode': 'Substring',
            //         'ContainmentComparison': 'IgnoreCase',
            //     },
            //     't:FieldURI': { attributes: { FieldURI: 'item:Subject' } },
            //     't:Constant': { attributes: { Value: subject } }
            // }
        }
        filterBody.subject = subject
    }
    if (to) {
        findMailEwsArgs.Restriction['t:And']['t:Contains'] = {
            'attributes': {
                'ContainmentMode': 'Substring',
                'ContainmentComparison': 'IgnoreCase',
            },
            't:FieldURI': { attributes: { FieldURI: 'item:DisplayTo' } },
            't:Constant': { attributes: { Value: to } }
        }
    }

    console.log('try fetching emails');
    // console.log(JSON.stringify(findMailEwsArgs, null, 2));

    let errorHandler = new promiseErrorHandler()

    return client.api(CONFIGS.graphPath)
        .filter(`ReceivedDateTime ge ${new Date(filterBody.start).toISOString()} and
             ReceivedDateTime le ${new Date(filterBody.end).toISOString()} and 
             contains(subject, '${filterBody.subject}')`)
        .select('body,subject')
        .top(5)
        .get()
        .catch(error => {
            return errorHandler.chain(error, `call fetching emails api fail`);
        })
        .then(result => {
            // console.log(`Success: ${JSON.stringify(result, null, 2)}`);

            console.log('[success] call fetching emails api success');

            if (!result.value) throw new Error('cannot find target email');

            let targets = result.value;

            let emails;
            // 可能多筆，可能一筆，回傳結構不同
            emails = targets.length ? targets.map(item => {
                return {
                    Subject: item.subject,
                    Id: item.id,
                    Body: item.body
                }
            }) : [{
                Subject: targets.subject,
                Id: targets.id,
                Body: targets.body
            }];

            console.log(`"${emails.length}" emails matche filter conditions:`);
            // console.log(emails.map(email => `------ ${email.Subject}`).join('\n'));

            if (emails.length == 0) {
                throw new Error('cannot find target email');
            } else if (emails.length > 1) {
                // throw new Error('more than one email matches filter conditions target email');
                console.log('more than one email matches filter conditions target email');
                emails.shift();
            }

            console.log('[success] fetching target email success');

            return Promise.resolve(emails[0]);
        })
        .catch(error => {
            return errorHandler.chain(error, `parsing email content fail`, true);
        })
}

// 新版作法不用了!
exports.getEmailContent = async (email) => {

    let getMailEwsArgs = {
        'ItemShape': {
            'BaseShape': 'IdOnly',
            'AdditionalProperties': {
                'FieldURI': [
                    { 'attributes': { 'FieldURI': 'item:Subject' } },
                    { 'attributes': { 'FieldURI': 'item:Body' } }
                ]
            }
        },
        'ItemIds': {
            'ItemId': {
                'attributes': {
                    'Id': email.Id
                }
            }
        }
    }

    let errorHandler = new promiseErrorHandler()

    return ews.run('GetItem', getMailEwsArgs)
        .catch(error => {
            return errorHandler.chain(error, `call fetching email content api fail`);
        })
        .then(response => {
            console.log('[success] call fetching email content api success');

            try {
                email.body = response.ResponseMessages.GetItemResponseMessage.Items.Message.Body.$value;
            }
            catch (error) {
                throw new Error(`not expected email object format`);
            }

            console.log('[success] get email content');

            return Promise.resolve(email);
        })
        .catch(error => {
            return errorHandler.chain(error, `fetching email content fail`, true);
        });

}

// 篩選內文，找出需推播事項
// 多看兩下
exports.filterContent = async (email) => {
    console.log(`start parsing "${email.Subject}" email content`);

    try {
        // console.log('email.body', email.Body.content);
        let htmlObj = HTMLPARSER.parse(email.Body.content);
        let tableObjs = htmlObj.querySelectorAll('table');
        let selectedMessages = [];

        tableObjs[5].querySelectorAll('tr').forEach((tr, i) => {
            let tds = tr.querySelectorAll('td');

            if (tds[0].text.trim().indexOf('Successful') < 0 && 0 != i) {
                let myMsg = `Policy Name : ${tds[3].text.trim()} , Error Code  ${tds[tds.length - 1].text.trim()}`;

                if (selectedMessages.indexOf(myMsg) == -1)
                    selectedMessages.push(myMsg);
            }
        });

        // 加入編號
        selectedMessages = selectedMessages.map((msg, i) => `${i + 1}. ${msg}`);

        email.errorCount = selectedMessages.length;
        email.messages = selectedMessages.join('\n');
        delete email.body;

        console.log(`[success] email "${email.Subject}" parsing success, ${selectedMessages.length} messages should be reported`);

        // console.log(JSON.stringify(email, null, 2));
        return Promise.resolve(email);

    } catch (error) {
        return Promise.reject(`[fail] fail to parse email, ${error.message}`);
    }
}

// 獨立推播動作
exports.pushToTarget = async (message, token, email, pushId) => {
    console.log('start pushing messages');

    let messages = [];

    // 有錯誤訊息
    if (email.errorCount) {
        console.log(`${email.messages.length} messages preparing to broadcast`);
        messages = [];
        let wholeMessage;

        // 錯誤訊息過多判定
        if (email.errorCount >= 50) {
            wholeMessage = `${email.Subject + '\n'}今日 (${new Date().toLocaleDateString('zh-Hans-CN')}) 備份失敗，其中 ${email.errorCount} 筆錯誤\n * * * * * \n 請檢查備份訊息。`
        } else {
            wholeMessage = `${message ? message + '\n' : ''}${email.Subject + '\n'}今日 (${new Date().toLocaleDateString('zh-Hans-CN')}) 備份失敗，其中 ${email.errorCount} 筆錯誤\n * * * * * \n${email.messages}`
        }
        // line 限制單筆最多 5000 字
        let limit = 4999;
        for (i = 0; i < wholeMessage.length; i += limit) {
            messages.push({
                type: 'text',
                text: wholeMessage.substr(i, limit * (i + 1))
            })
        }
        // 無錯誤訊息
    } else if (email.Subject === 'NBU-Savh-Daily Report' ||
        email.Subject === 'NBU-nbuys-Daily Report') {
        // 蘇澳榮院 員山榮院 格式
        messages = [{
            type: 'text',
            text: `${message ? message + '\n' : ''}${email.Subject + '\n'}今日 (${new Date().toLocaleDateString('zh-Hans-CN')}) 本地備份與異地AIR全部成功`
        }]
    }
    else {
        // console.log(`1 筆系統正常訊息要推播`);
        messages = [{
            type: 'text',
            text: `${message ? message + '\n' : ''}${email.Subject + '\n'}今日 (${new Date().toLocaleDateString('zh-Hans-CN')}) 備份全部成功`
        }]
    }

    // line限制五個訊息一筆

    let messagesArray = [];
    for (i = 0; i < messages.length; i += 5) {
        messagesArray.push(messages.slice(i, i + 5));
    }
    let pushRequests = [];
    for (i = 0; i < messagesArray.length; i++) {
        pushRequests.push(pushToLine(token, {
            to: pushId,
            messages: messagesArray[i]
        }))
    }
    return timeoutPromiseTrain(pushRequests)
        .then(response => {
            return Promise.resolve();
        })
        .catch(error => {
            return Promise.reject(error);
        })
        .finally(() => {
            console.log('finish push tasks');
        })
}

// 群播動作
exports.broadcastToMembers = async (token, email) => {
    console.log('start to broadcast messages');

    let messages = [];

    // 有錯誤訊息
    if (email.messages.length) {
        console.log(`${email.messages.length} messages preparing to broadcast`);
        messages = email.messages.map(msg => {
            return {
                type: 'text',
                text: `今日 (${new Date().toLocaleDateString('zh-Hans-CN')}) 備份失敗，其中 ${email.errorCount} 筆錯誤\n * * * * * \n${msg}`
            }
        })

        // 無錯誤訊息
    } else {
        // console.log(`1 筆系統正常訊息要推播`);
        messages = [{
            type: 'text',
            text: `今日 (${new Date().toLocaleDateString('zh-Hans-CN')}) 備份全部成功`
        }]
    }

    // line限制五個訊息一筆
    let messagesArray = [];
    for (i = 0; i < messages.length; i += 5) {
        messagesArray.push(messages.slice(i, i + 5));
    }

    let broadcastRequests = [];
    for (i = 0; i < messagesArray.length; i++) {
        broadcastRequests.push(broadcastLine(token, {
            messages: messagesArray[i]
        }))
    }

    return timeoutPromiseTrain(broadcastRequests)
        .then(response => {
            return Promise.resolve();
        })
        .catch(error => {
            return Promise.reject(error);
        })
        .finally(() => {
            console.log('finish broadcast tasks');
        })
}

// LINE 推播與失敗重試
function broadcastLine(token, params, retry = 0) {
    let endPoint = {
        'url': 'https://api.line.me/v2/bot/message/broadcast',
        'method': 'POST',
        'headers': {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
        },
        'body': JSON.stringify(params)
    };

    console.log(`start broadcasting, ${retry++}st try`);

    return request(endPoint)
        .then(response => {
            console.log(`[success] broadcast success`);
            return Promise.resolve();
        })
        .catch(error => {
            if (retry < CONFIGS.pushRetry) {
                console.log(`[fail] broadcast fail, ${retry}st try`);
                return new Promise((resolve, reject) => {
                    setTimeout(() => {
                        resolve(pushToLine(token, params, retry));
                    }, CONFIGS.pushTimeout);
                });
            } else {
                console.log(`[fail] broadcast fail, exceed limit retry times (${CONFIGS.pushRetry}), give up!`);
                return Promise.resolve();
            }
        });
}
// LINE 單獨發送與失敗重試
function pushToLine(token, params, retry = 0) {
    let endPoint = {
        'url': 'https://api.line.me/v2/bot/message/push',
        'method': 'POST',
        'headers': {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
        },
        'body': JSON.stringify(params)
    };

    console.log(`pushing message to ${params.to}, ${retry++}st try`);

    // return new Promise((resolve, reject) => {
    //     console.log(JSON.stringify(params));
    //     resolve();
    // });

    return request(endPoint)
        .then(response => {
            console.log(`[success] pushing messages success`);
            return Promise.resolve();
        })
        .catch(error => {
            if (retry < CONFIGS.pushRetry) {
                console.log(`[fail] pushing messages fail, ${retry}st fail`);
                return new Promise((resolve, reject) => {
                    setTimeout(() => {
                        resolve(pushToLine(token, params, retry));
                    }, CONFIGS.pushTimeout);
                });
            } else {
                console.log(`[fail] push message to ${params.to} fail, exceed limit retry times (${CONFIGS.pushRetry}), give up!`);
                return Promise.resolve();
            }
        });
}
// 排列並延時發Request
function timeoutPromiseTrain(funcs) {
    if (funcs.length == 0) return Promise.resolve();

    let subFuncs = funcs.slice(0, CONFIGS.pushRequestOnce);

    return new Promise((resolve, reject) => {
        setTimeout(() => {
            Promise.all(subFuncs)
                .then(() => {
                    funcs.splice(0, CONFIGS.pushRequestOnce);
                    resolve(timeoutPromiseTrain(funcs));

                }).catch(error => {
                    console.log('unexpected error', error);
                    reject();
                })
        }, CONFIGS.pushTimeout);
    });
}
// 台灣UTC字串
function taiwanUTCString(taiwanTime) {
    let timeZoneOffset = new Date().getTimezoneOffset() / 60;
    return new Date(new Date(taiwanTime).getTime() + ((-8 - timeZoneOffset) * 60 * 60 * 1000)).toISOString();
}
// 錯誤管理器
class promiseErrorHandler {
    constructor() {
        this.errMsg;
    }
    chain(error, msg, final) {
        if (!this.errMsg) {
            this.errMsg = msg;
        }
        if (final) {
            console.error('[fail]', this.errMsg, error);
            return Promise.reject(this.errMsg);
        } else {
            return Promise.reject(error);
        }
    }
};