/*-----------------------------------------------------------------------------
An Office 365 Language Understanding (LUIS) bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");
var AuthenticationContext = require('adal-node').AuthenticationContext;
var Promise = require('es6-promise').Promise;
var request = require("request");
var util = require('util');
var moment = require('moment');
var fs  = require('fs');
var PromiseX = require('bluebird');
var requestX = require('request-promise').defaults({ encoding: null });

//=========================================================
// ADAL Configuration
//=========================================================
var adalConfig = {
    'clientId': process.env.AAD_CLIENT_ID, // The client Id retrieved from the Azure AD App
    'clientSecret': process.env.AAD_CLIENT_SECRET, // The client secret retrieved from the Azure AD App
    'authorityHostUrl': 'https://login.microsoftonline.com/', // The host URL for the Microsoft authorization server
    'tenant': process.env.TENANT, // The tenant Id or domain name (e.g mydomain.onmicrosoft.com)
    'redirectUri': process.env.REDIRECT_URI, // This URL will be used for the Azure AD Application to send the authorization code.
    'resource': process.env.RESOURCE, // The resource endpoint we want to give access to (in this case, SharePoint Online)
    'state': process.env.STATE,
}

adalConfig.authorityUrl = adalConfig.authorityHostUrl + adalConfig.tenant;
adalConfig.templateAuthzUrl = adalConfig.authorityUrl +
    '/oauth2/authorize?response_type=code&client_id=' + // Optionally, we can get an Open Id Connect id_token to get more info on the user (some additional parameters are required if so https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-openid-connect-code)
    adalConfig.clientId +
    '&state=<state>' +
    '&response_mode=form_post' + //We want response as POST http request (see callback to see why)
    '&redirect_uri=' + adalConfig.redirectUri +
    '&resource=' + adalConfig.resource;

// Setup Restify Server
var server = restify.createServer();

server.use(restify.plugins.bodyParser({ // To be able to get the authorization code (req.params.code)
    mapParams: true
}));

// Graph
var graph = {};

server.listen(process.env.port || process.env.PORT || 3978, function () {
    console.log('%s listening to %s', server.name, server.url);
});

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

// Create a route for the Azure AD App callback
// Be careful here: if you specify a GET request for the OAuth callback, IISNODE will interpret the response as a static file due to the query string parameters instead of redirect it to the correct node js server route.
// To avoid modify the web.config, use a POST request instead
server.post('/api/oauthcallback', (req, res, next) => {

    // Get the authorization code from the Azure AD application
    var authorizationCode = req.params.code;
    if (authorizationCode) {

        acquireTokenWithAuthorizationCode(authorizationCode).then((response) => {

            // Add the state to the response to validate the CSRF scenario
            // The state has two utilities here:
            // - Reconnect with the bot to continue dialog
            // - Avoid CRSF attacks
            var state = req.params.state;
            if (state) {

                var address = JSON.parse(state);
                response.state = state;

                // Continue the dialog with the bot. Be careful, beginDialog" starts a new conversation.
                // We use the state parameter to save the address and be able to reconnect with the bot after authentication
                // Special thanks to this blog post https://dev-hope.blogspot.ca/2016/09/google-oauth-using-nodejs-and-microsoft.html
                // https://docs.botframework.com/en-us/node/builder/chat/UniversalBot/#navtitle ==> See paragraph "Saving Users Address"
                bot.beginDialog(address, "oauth-success", response);
            }

            var body = '<html><body>Authentication succeeded! You can now close this tab</body></html>';
            res.send(200, body, { 'Content-Length': Buffer.byteLength(body), 'Content-Type': 'text/html' });
            res.end();

        }).catch((errorMessage) => {

            var body = '<html><body>' + errorMessage + '</body></html>';
            res.send(200, body, { 'Content-Length': Buffer.byteLength(body), 'Content-Type': 'text/html' });
            res.end();
        });

    } else {

        var body = '<html><body>Something went wrong, we didn\'t get an authorization code</body></html>';
        res.send(200, body, { 'Content-Length': Buffer.byteLength(body), 'Content-Type': 'text/html' });
        res.end();
    }
});

//=========================================================
// Bot authorization delegation middleware
//=========================================================
var getAuthorization = async (session, args) => {
    // User is not already signed-in
    if (!session.privateConversationData['accessToken']) {
        // Set the arbitrary state as the current session address
        var stateToken = encodeURIComponent(JSON.stringify(session.message.address))
        var authorizationUrl = adalConfig.templateAuthzUrl.replace('<state>', stateToken);
        var actionLabel = 'You need to sign in to Office 365 before to start!';
        var buttonLabel = 'Sign-in';
        var signInCard = null;
        // The Sign-In card is not supported by Microsoft Teams for now (23/01/2017)
        // https://msdn.microsoft.com/en-us/microsoft-teams/bots#cards-and-buttons
        if (session.message.address.channelId === "msteams") {
            var link = builder.CardAction.openUrl(session, authorizationUrl, buttonLabel)
            signInCard = new builder.ThumbnailCard(session)
                .title("Authorization required!")
                .text(actionLabel)
                .buttons([link]);
        } else {
            // Send sign-in card
            signInCard = new builder.SigninCard(session)
                .text(actionLabel)
                .button(buttonLabel, authorizationUrl);
        }
        var msg = new builder.Message(session).attachments([signInCard]);
        session.send(msg);
    } else {
        // If the user is  already signed-in, we check if the access token is expired
        var expiresOn = session.privateConversationData['expiresOn'];
        var refreshToken = session.privateConversationData['refreshToken']
        var currentDateUTC = moment.utc().format('YYYY-MM-DDTHH:mm:ss.SSS[Z]');
        if (moment(currentDateUTC).isAfter(expiresOn)) {
            var response = await acquireTokenWithRefreshToken(refreshToken);
            // Refresh the token infos
            session.privateConversationData['accessToken'] = response.accessToken;
            session.privateConversationData['expiresOn'] = response.expiresOn;
            session.privateConversationData['refreshToken'] = response.refreshToken;
        }
    }
}

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

var tableName = 'botdata';
var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

// Create your bot with a function to receive messages from the user
// This default message handler is invoked if the user's utterance doesn't
// match any intents handled by other dialogs.
var bot = new builder.UniversalBot(connector, function (session, args) {
    session.send('You reached the default message handler. You said \'%s\'.', session.message.text);
    // Get back to the main dialog route and prompt for a sign in
    //session.beginDialog("Signin");
});

bot.set('storage', tableStorage);

// Make sure you add code to validate these fields
var luisAppId = process.env.LuisAppId;
var luisAPIKey = process.env.LuisAPIKey;
var luisAPIHostName = process.env.LuisAPIHostName || 'westus.api.cognitive.microsoft.com';

const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v2.0/apps/' + luisAppId + '?subscription-key=' + luisAPIKey;

// Create a recognizer that gets intents from LUIS, and add it to the bot
var recognizer = new builder.LuisRecognizer(LuisModelUrl);
bot.recognizer(recognizer);

//=========================================================
// Bot Dialogs
//=========================================================
bot.dialog('oauth-success', function (session, response) {
    // Check the state value to avoid CSRF attacks http://www.twobotechnologies.com/blog/2014/02/importance-of-state-in-oauth2.html
    if (encodeURIComponent(JSON.stringify(session.message.address)) !== encodeURIComponent(response.state)) {
        session.send("CSRF scenario detected. Closing the current conversation...");
        session.endDialog();
    } else {
        // Save the token for the current user and for this conversation only (privateConversationData)
        if (!session.privateConversationData['accessToken']) {

            session.privateConversationData['accessToken'] = response.accessToken;
            session.privateConversationData['expiresOn'] = response.expiresOn;
            session.privateConversationData['refreshToken'] = response.refreshToken;
        }
        session.send('Hi %s. What can I do for you today?', response.userName);
        // Get back to the main dialog route
        //session.beginDialog("/");
    }
});

// Add a dialog for each intent that the LUIS app recognizes.
// See https://docs.microsoft.com/en-us/bot-framework/nodejs/bot-builder-nodejs-recognize-intent-luis 
bot.dialog('GreetingDialog',
    (session) => {
        session.send('You reached the Greeting intent. You said \'%s\'.', session.message.text);
        session.endDialog();
    }
).triggerAction({
    matches: 'Greeting'
})

bot.dialog('HelpDialog',
    (session) => {
        session.send('You reached the Help intent. You said \'%s\'.', session.message.text);
        session.endDialog();
    }
).triggerAction({
    matches: 'Help'
})

bot.dialog('CancelDialog',
    (session) => {
        session.send('You reached the Cancel intent. You said \'%s\'.', session.message.text);
        session.endDialog();
    }
).triggerAction({
    matches: 'Cancel'
})

bot.dialog('Welcome',
    (session) => {
        session.send('Hi Giuliano');
        session.endDialog();
    }
);

bot.dialog('Signin', [getAuthorization,
    (session) => {

        var keywords = session.message.text

        // Check if a a message has been typed
        if (keywords) {

            // For debugging purpose, we add an arbitrary command to reset the bot state (we also could have implement a logout mechanism).
            // Initially the native /deleteprofile command was used but it is not available in the Bot Framework v3 anymore.
            if (keywords === "reset") {
                session.privateConversationData = {};

                // Get back to the main dialog route and prompt for a sign in
                session.beginDialog("SignIn");
            } else {

                session.beginDialog("Welcome");
            }
        }
    }]
).triggerAction({
    matches: 'Signin'
})

bot.dialog('ManagerDialog',
    async (session) => {
        await getAuthorization(session);
        if (session.privateConversationData['accessToken']) {
            graph.getMyManager(session.privateConversationData['accessToken'])
                .then((result) => {
                    console.log(result);
                    session.send('Your manager is \'%s\'.', result);
                    session.endDialog();
                }, (error) => {
                    console.error('>>> Error getting manager: ' + error.message);
                    session.send('>>> Error getting manager: %s' + error.message);
                    session.endDialog();
                });
        }
    }
).triggerAction({
    matches: 'Manager'
})

bot.dialog('PhotoDialog',
    async (session) => {
        await getAuthorization(session);
        graph.getMyPhoto(session.privateConversationData['accessToken'])
            .then((result) => {
                console.log(result);
                var attachmentFileName = 'myPicture.jpg';
                var base64 = Buffer.from(result).toString('base64');
                var contentType = 'image/jpg';
                //var test = 'data:image/jpeg;base64,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';
                var msg = new builder.Message(session)
                    .addAttachment({
                        contentUrl: util.format('data:%s;base64,%s', contentType, base64),
                        contentType: contentType,
                        name: attachmentFileName
                    });

                session.send(msg);
                session.endDialog();
            }, (error) => {
                console.error('>>> Error getting my photo: ' + error.message);
                session.send('>>> Error getting my photo: %s' + error.message);
                session.endDialog();
            });
    }
).triggerAction({
    matches: 'Photo'
})

// Dialog to change the profile picture
bot.dialog('ChangePhotoDialog', [
    (session) => {
        var msg = new builder.Message(session)
            .speak('Please upload your new picture')
            .text('Please upload your new picture');
        builder.Prompts.attachment(session, msg);
    },
    async (session, results) => {
        await getAuthorization(session);
        var msg = session.message;
        if (msg.attachments.length) {
            // Message with attachment, proceed to download it.
            // Skype & MS Teams attachment URLs are secured by a JwtToken, so we need to pass the token from our bot.
            var attachment = msg.attachments[0];
            var fileDownload = checkRequiresToken(msg)
                ? requestWithToken(attachment.contentUrl)
                : requestX(attachment.contentUrl);
    
            fileDownload.then(
                function (picture) {
                    graph.setMyPhoto(session.privateConversationData['accessToken'], picture).then((response) => {
                        //session.endDialogWithResult({ response: results.response });
                        builder.Prompts.text(session, 'Your picture has been changed successfully.');
                        session.endDialog();
                    });
                }).catch(function (err) {
                    console.log('Error downloading attachment:', { statusCode: err.statusCode, message: err.response.statusMessage });
                });
        } else {
            // No attachments were sent
            var reply = new builder.Message(session)
                .text('No attachment was sent to me. Please try again sending a new message with an attachment.');
            session.send(reply);
            // Repeat the dialog
            session.replaceDialog('ChangePhotoDialog', { reprompt: true });
        }
    }
]).triggerAction({
    matches: 'ChangePhoto'
});

//=========================================================
// Helper Methods
//=========================================================

// Request file with Authentication Header
var requestWithToken = function (url) {
    return obtainToken().then(function (token) {
        return request({
            url: url,
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/octet-stream'
            }
        });
    });
};

// Promise for obtaining JWT Token (requested once)
var obtainToken = PromiseX.promisify(connector.getAccessToken.bind(connector));

var checkRequiresToken = function (message) {
    return message.source === 'skype' || message.source === 'msteams';
};

//=========================================================
// ADAL Helper Methods
//=========================================================
var acquireTokenWithAuthorizationCode = (authorizationCode) => {
    var authenticationContext = new AuthenticationContext(adalConfig.authorityUrl);
    var p = new Promise((resolve, reject) => {
        authenticationContext.acquireTokenWithAuthorizationCode(
            authorizationCode,
            adalConfig.redirectUri, // This URL must be the same as the redirect_uri of the original request or the reply url of the Azure AD App. Otherwise, it will throw an error.
            adalConfig.resource,
            adalConfig.clientId,
            adalConfig.clientSecret,
            (err, response) => {

                if (err) {
                    reject('error: ' + err.message + '\n');

                } else {
                    resolve({
                        userName: (response.givenName + " " + response.familyName),
                        accessToken: response.accessToken,
                        expiresOn: response.expiresOn,
                        refreshToken: response.refreshToken,
                    });
                }
            });
    });
    return p;
}

var acquireTokenWithRefreshToken = (refreshToken) => {
    var authenticationContext = new AuthenticationContext(adalConfig.authorityUrl);
    var p = new Promise((resolve, reject) => {
        authenticationContext.acquireTokenWithRefreshToken(
            refreshToken,
            adalConfig.clientId,
            adalConfig.clientSecret,
            adalConfig.resource,
            (err, response) => {
                if (err) {
                    reject(errorMessage = 'error: ' + err.message + '\n');

                } else {
                    resolve({
                        userName: (response.givenName + " " + response.familyName),
                        accessToken: response.accessToken,
                        expiresOn: response.expiresOn,
                        refreshToken: response.refreshToken,
                    });
                }
            });
    });
    return p;
}

//=========================================================
// GRAPH API Helper Methods
//=========================================================

/**
 * Get my manager
 * @param {*} token to append in the header in order to make the request
 */
graph.getMyManager = (token) => {
    var p = new Promise((resolve, reject) => {
        request.get('https://graph.microsoft.com/v1.0/me/manager', {
            auth: {
                bearer: token
            }
        }, function (err, response, body) {
            var parsedBody = JSON.parse(body);

            if (err) {
                reject(err);
            } else if (parsedBody.error) {
                reject(parsedBody.error.message);
            } else {
                // The value of the body will be an array.
                resolve(parsedBody.displayName);
            }
        });
    });
    return p;
};

/**
 * Get my profile's photo
 * @param {*} token to append in the header in order to make the request
 */
graph.getMyPhoto = (token) => {
    var p = new Promise((resolve, reject) => {
        request.get('https://graph.microsoft.com/v1.0/me/photo/$value', {
            auth: {
                bearer: token
            }, encoding: null,
        }, function (err, response, body) {
            if (err) {
                reject(err);
            } else if (body.error) {
                reject(body.error.message);
            } else {
                // The value of the body will be an array.
                resolve(body);
            }
        });
    });
    return p;
};

/**
 * Set my profile's photo
 * @param {*} token to append in the header in order to make the request
 */
graph.setMyPhoto = (token, picture) => {
    var p = new Promise((resolve, reject) => {
        request.patch('https://graph.microsoft.com/v1.0/me/photo/$value', {
            auth: {
                bearer: token
            },
            headers: {
                'Content-type': 'image/jpeg'
            },
            body: picture
        }, function (err, response, body) {
            if (err) {
                reject(err);
            } else if (body.error) {
                reject(body.error.message);
            } else {
                // The value of the body will be an array.
                resolve(body);
            }
        });
    });
    return p;
};
