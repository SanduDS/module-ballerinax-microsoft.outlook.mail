import ballerina/log;
import ballerina/os;
import ballerinax/microsoft.outlook.mail;

configurable string & readonly refreshUrl = os:getEnv("TOKEN_ENDPOINT");
configurable string & readonly refreshToken = os:getEnv("REFRESH_TOKEN");
configurable string & readonly clientId = os:getEnv("APP_ID");
configurable string & readonly clientSecret = os:getEnv("APP_SECRET");

mail:Configuration configuration = {clientConfig: {
        refreshUrl: refreshUrl,
        refreshToken: refreshToken,
        clientId: clientId,
        clientSecret: clientSecret
    }};

mail:OutlookClient oneDriveClient = check new (configuration);

public function main() returns error? {
    var result = check oneDriveClient->listAttachment("<Message ID>", "<Folder ID or Well Known Name");
    int index = 0;
    error? e = result.forEach(function(mail:FileAttachment queryResult) {
                                  index += 1;
                              });
    log:printInfo("Total Count of  Attachments : " + index.toString());
}
