import ballerina/log;
import ballerina/os;
import ballerinax/microsoft.outlook.mail;

configurable string & readonly refreshUrl = os:getEnv("TOKEN_ENDPOINT");
configurable string & readonly refreshToken = os:getEnv("REFRESH_TOKEN");
configurable string & readonly clientId = os:getEnv("APP_ID");
configurable string & readonly clientSecret = os:getEnv("APP_SECRET");

mail:Configuration configuration = {
    clientConfig: {
        refreshUrl: refreshUrl,
        refreshToken : refreshToken,
        clientId : clientId,
        clientSecret : clientSecret
    }
};

mail:OutlookClient oneDriveClient = check new(configuration);

public function main() returns error? {
    mail:DraftMessage draft = {
        subject:"<Mail Subject>",
        importance:"Low",
        body:{
            "contentType":"HTML",
            "content":"We are <b>Wso2</b>!"
        },
        toRecipients:[
            {
                emailAddress:{
                    address:"<Your Email Address>",
                    name: "Name (Optional)"
                }
            }
        ]
    };
    mail:Message message = check oneDriveClient->createMessage(draft);
    log:printInfo(message.toString());
}
