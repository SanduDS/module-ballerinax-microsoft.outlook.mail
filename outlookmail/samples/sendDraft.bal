import ballerina/log;
import ballerina/os;
import dhanushkas/microsoft.outlook.mail;

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
    error? output = oneDriveClient->sendDraftMessage("<createdDraftId>");
    if (output is error) {
        log:printError(msg = "Sending Draft Failed",'error = output);
    }
}
