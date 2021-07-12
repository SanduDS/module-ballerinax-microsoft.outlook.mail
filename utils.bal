import ballerina/http;

isolated function sendRequestGET(http:Client httpClient, string resources) returns @tainted json|error {
    return httpClient->get(resources, targetType=json);
}

isolated function getOutlookClient(Configuration config, string baseUrl) returns http:Client|error {
    http:BearerTokenConfig|http:OAuth2RefreshTokenGrantConfig clientConfig = config.clientConfig;
    http:ClientSecureSocket? socketConfig = config?.secureSocketConfig;
    return check new (baseUrl, {
            auth: clientConfig,
            secureSocket: socketConfig
    });
}

isolated  function getRecipientListAsRecord(string comment, string[] addressList) returns ForwardParamsList {
    Recipient[] recipients = [];
    foreach string address in addressList {
        EmailAddress emailAddress = {
            address: address
        };
        Recipient recipient = {
            emailAddress :emailAddress
        };
        recipients.push(recipient);
    }
    return {comment: comment, toRecipients: recipients};
}
