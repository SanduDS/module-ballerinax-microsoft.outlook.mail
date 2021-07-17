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

public function main() {
  log:printInfo("oneDriveClient->listMessages()");
    var output = oneDriveClient->listMessages(folderId = "drafts", optionalUriParameters="?$select:\"sender,subject\"");
    if (output is stream<mail:Message, error?>) {
        int index = 0;
        error? e = output.forEach(function (mail:Message queryResult) {
            index = index + 1; 
            log:printInfo(queryResult?.id.toString());
        });
        log:printInfo("Total count of records : " +  index.toString());        
    } else {
        log:printError(output.toString());
    }  
}
