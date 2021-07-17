// import ballerina/log;
// import ballerina/os;
// import dhanushkas/microsoft.outlook.mail;
// configurable string & readonly refreshUrl = os:getEnv("TOKEN_ENDPOINT");
// configurable string & readonly refreshToken = os:getEnv("REFRESH_TOKEN");
// configurable string & readonly clientId = os:getEnv("APP_ID");
// configurable string & readonly clientSecret = os:getEnv("APP_SECRET");
// mail:Configuration configuration = {
//     clientConfig: {
//         refreshUrl: refreshUrl,
//         refreshToken : refreshToken,
//         clientId : clientId,
//         clientSecret : clientSecret
//     }
// };
// mail:OutlookClient oneDriveClient = check new(configuration);
// public function main() {
//   log:printInfo("oneDriveClient->listMessages()");
//     var output = oneDriveClient->listMessages(folderId = "drafts", optionalUriParameters="?$select:\"sender,subject\"");
//     if (output is stream<mail:Message, error?>) {
//         int index = 0;
//         error? e = output.forEach(function (mail:Message queryResult) {
//             index = index + 1; 
//             log:printInfo(queryResult?.id.toString());
//         });
//         log:printInfo("Total count of records : " +  index.toString());        
//     } else {
//         log:printError(output.toString());
//     }  
// }
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// import ballerina/log;
// import ballerina/os;
// import dhanushkas/microsoft.outlook.mail;
// configurable string & readonly refreshUrl = os:getEnv("TOKEN_ENDPOINT");
// configurable string & readonly refreshToken = os:getEnv("REFRESH_TOKEN");
// configurable string & readonly clientId = os:getEnv("APP_ID");
// configurable string & readonly clientSecret = os:getEnv("APP_SECRET");
// mail:Configuration configuration = {
//     clientConfig: {
//         refreshUrl: refreshUrl,
//         refreshToken : refreshToken,
//         clientId : clientId,
//         clientSecret : clientSecret
//     }
// };
// mail:OutlookClient oneDriveClient = check new(configuration);
// public function main() returns error? {
//     mail:DraftMessage draft = {
//         subject:"<Mail Subject>",
//         importance:"Low",
//         body:{
//             "contentType":"HTML",
//             "content":"We are <b>Wso2</b>!"
//         },
//         toRecipients:[
//             {
//                 emailAddress:{
//                     address:"<Your Email Address>",
//                     name: "Name (Optional)"
//                 }
//             }
//         ]
//     };
//     mail:Message message = check oneDriveClient->createMessage(draft);
//     log:printInfo(message.toString());
// }
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// import ballerina/log;
// import ballerina/os;
// import dhanushkas/microsoft.outlook.mail;
// configurable string & readonly refreshUrl = os:getEnv("TOKEN_ENDPOINT");
// configurable string & readonly refreshToken = os:getEnv("REFRESH_TOKEN");
// configurable string & readonly clientId = os:getEnv("APP_ID");
// configurable string & readonly clientSecret = os:getEnv("APP_SECRET");
// mail:Configuration configuration = {clientConfig: {
//         refreshUrl: refreshUrl,
//         refreshToken: refreshToken,
//         clientId: clientId,
//         clientSecret: clientSecret
//     }};
// mail:OutlookClient oneDriveClient = check new (configuration);
// public function main() returns error? {
//     var result = check oneDriveClient->listAttachment("<sentMessageId>", "sentitems");
//     int index = 0;
//     error? e = result.forEach(function(mail:FileAttachment queryResult) {
//                                   index += 1;
//                               });
//     log:printInfo("Total Count of  Attachments : " + index.toString());
// }
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// import ballerina/log;
// import ballerina/os;
// import dhanushkas/microsoft.outlook.mail;
// configurable string & readonly refreshUrl = os:getEnv("TOKEN_ENDPOINT");
// configurable string & readonly refreshToken = os:getEnv("REFRESH_TOKEN");
// configurable string & readonly clientId = os:getEnv("APP_ID");
// configurable string & readonly clientSecret = os:getEnv("APP_SECRET");
// mail:Configuration configuration = {clientConfig: {
//         refreshUrl: refreshUrl,
//         refreshToken: refreshToken,
//         clientId: clientId,
//         clientSecret: clientSecret
//     }};
// mail:OutlookClient oneDriveClient = check new (configuration);
// public function main() returns error? {
//     var result = check oneDriveClient->listMessages("<Folder ID>", 
//     optionalUriParameters = "?$select: \"sender,subject,hasAttachments\"");
//     int index = 0;
//     error? e = result.forEach(function(mail:Message queryResult) {
//                                   index += 1;
//                               });
//     log:printInfo("Total Count of  Attachments : " + index.toString());
// }
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
    mail:FileAttachment attachment = {
        contentBytes: "SGVsbG8gV29ybGQh",
        contentType: "text/plain",
        name: "sample.txt"
    };
    mail:FileAttachment fileAttachment = check oneDriveClient->addFileAttachment("<Message ID>", attachment, 
    "<Folder ID>");
    log:printInfo(fileAttachment.toString());
}
