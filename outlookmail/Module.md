## Overview
Ballerina connector for Microsoft Outlook mail provides access to Microsoft Outlook mail service in Microsoft Graph v1.0 via Ballerina language easily. It provides capability to perform more useful functionalities provided in Microsoft outlook mail such as sending messages, listing messages, creating drafts, mail folders, deleting and updating messages etc. 

This module supports Ballerina SL Beta 2 version.
 
## Configuring connector
### Prerequisites
* Microsoft Outlook Account
* Azure Active Directory Access
### Obtaining tokens
Please follow [this link](https://docs.microsoft.com/en-us/graph/auth-v2-user#authentication-and-authorization-steps) and obtain the client ID, client secret and refresh token.
 
## Quickstart
* Send a message with an attachment
Step 1: Import MS Outlook Mail Package
First, import the ballerinax/microsoft.outlook.mail module into the Ballerina project.
```ballerina
import ballerinax/microsoft.outlook.mail;
```
Step 2: Configure the connection to an existing Azure AD app
You can now make the connection configuration using the OAuth2 refresh token grant config.
```ballerina
outlookMail:Configuration configuration = {
    clientConfig: {
        refreshUrl: <REFRESH_URL>,
        refreshToken : <REFRESH_TOKEN>,
        clientId : <CLIENT_ID>,
        clientSecret : <CLIENT_SECRET>
    }
};

outlookMail:Client teamsClient = check new(configuration);

```
Step 3: send a message
```
function testSendMessage() {
    log:printInfo("oneDriveClient->testSendMessage()");
    FileAttachment attachment = {
        contentBytes: "SGVsbG8gV29ybGQh",
        contentType: "text/plain",
        name: "sample.txt"
    };
    MessageContent messageRequest = {
       message:  {
            subject:"Ballerina Outlook Connector",
            importance:"Low",
            body:{
                "contentType":"HTML",
                "content":"<b>This is Ballerina</b>!"
            },
            toRecipients:[
                {
                    emailAddress:{
                        address:"sample@wso2.com",
                        name: "Sample"
                    }
                }
            ],
            attachments: [attachment]
        },
        saveToSentItems: true
    };
    
    var output = oneDriveClient->sendMessage(messageReq);
    if (output is error) {
        log:printError(output.toString());
    } else {
        log:printInfo(output.toString());
    }

``` 
## [You can find more samples here](https://github.com/ballerina-platform/module-ballerinax-microsoft.teams/tree/main/teams/samples)