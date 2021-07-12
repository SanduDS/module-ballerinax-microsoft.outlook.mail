// Copyright (c) 2021 WSO2 Inc. (http://www.wso2.org) All Rights Reserved.
//
// WSO2 Inc. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/http;
import ballerina/log;
import ballerina/lang.'int;
import ballerina/io;

 
# Ballerina Client for microsoft outlook mail operations
@display {
    label: "Microsoft Outlook.mail Client",
    iconPath: "MSOutlookMailLogo.svg"
}
public client class OutlookClient {
    http:Client httpClient;
    Configuration config;

    # Client Initialization
    #
    # + config - Configuration for client connector
    # + return - If success returns null otherwise returns the relevant error  
    public isolated function init(Configuration config) returns error? {
        self.httpClient = check getOutlookClient(config, BASE_URL);
        self.config = config;
    }

    # Get the messages in the signed-in user's mailbox (including the Deleted Items and Clutter folders)
    #
    # + folderId - The ID of the specific folder in the user's mailbox or the name of well-known folders (inbox, 
    # sentitems etc.)
    # + optionalUriParameters - The optional query parameter string
    # (https://docs.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http#optional-query-parameters)
    # + return - If success returns a ballerina stream of message records otherwise the relevant error
    @display {label: "List Messages"}  
    isolated remote function listMessages(@display {label: "Folder ID"} string? folderId = (), 
                                          @display {label: "Optional Query Parameters"} string? optionalUriParameters 
                                          = ()) returns @tainted error|stream<Message, error?> {
        string requestParams = optionalUriParameters is () ? "" : optionalUriParameters;
        requestParams = folderId is string ? "/mailFolders/" + folderId + "/messages" + requestParams : "/messages" + 
        requestParams;
        json response = check self.httpClient->get(requestParams, targetType = json);
        MessageStream objectInstance = check new (response, self.config);
        stream<Message, error?> finalStream = new (objectInstance);
        return finalStream;
    }

    # Create a draft of a new message
    #
    # + message - The detail of the draft  
    # + folderId - The mail folder where the draft should be saved in
    # + return - If success returns the newly created draft detail as a message record otherwise the relevant error
    @display {label: "Create Message"}
    isolated remote function createMessage(@display {label: "Draft Message"} DraftMessage message,
                                           @display {label: "Folder ID"} string? folderId = ()) returns @tainted 
                                           Message|error {
        string requestParams = folderId is string ? "/mailFolders/" + folderId + "/messages" : "/messages";
        http:Request request = new;
        request.setJsonPayload(message.toJson());
        request.setHeader("Content-Length", "0");
        request.setHeader("Content-Type", "application/json");
        return check self.httpClient->post(requestParams, request, targetType = Message);
    }

    # Retrieve the properties and relationships of a message
    #
    # + messageId - The ID of the message  
    # + folderId - The ID of the folder where the message is in
    # + optionalUriParameters - The OData Query Parameters to help customize the response
    # (https://docs.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0&tabs=http#optional-query-parameters)
    # + bodyContentType - The format of the body and uniqueBody of the specified message (values : html (default), text) 
    # + return - If success returns the requested message detail as a message record otherwise the relevant error
    @display {label: "Get Message"}
    isolated remote function getMessage(@display {label: "Message ID"} string messageId,
                                        @display {label: "Folder ID"} string? folderId = (), 
                                        @display {label: "Optional Query Parameters"} string? optionalUriParameters = 
                                        (), @display {label: "Body Content Format"} string? bodyContentType = ()) 
                                        returns 
                                        @tainted Message|error {
        string optionalPrams = optionalUriParameters is () ? "" : optionalUriParameters;
        string requestParams = folderId is string ? "/mailFolders/" + folderId : "";
        requestParams += "/messages/" + messageId + "/" + optionalPrams;
        if (bodyContentType is string) {
            map<string> headers = {"Prefer": "outlook.body-content-type=" + bodyContentType};
            return check self.httpClient->get(requestParams, headers, Message);
        }
        return check self.httpClient->get(path = requestParams, targetType = Message);
    }

    # Update the properties of a message 
    #
    # + messageId - The ID of the message  
    # + message - The message properties to be updated 
    # + folderId - The ID of the folder where the message is in
    # + return - If success returns the updated message detail as a message record otherwise the relevant error
    @display {label: "Update Message"}
    isolated remote function updateMessage(@display {label: "Update Message"} string messageId, 
                                           @display {label: "Message Content"}MessageUpdateContent message, 
                                           @display {label: "Folder ID"} string? folderId = ()) 
                                            returns @tainted Message|error {
        string requestParams = folderId is string ? "/mailFolders/" + folderId + "/messages/" : "/messages/";
        requestParams += messageId;
        http:Request request = new;
        request.setJsonPayload(message.toJson());
        request.setHeader("Content-Type", "application/json");
        return check self.httpClient->patch(requestParams, request, targetType = Message);
    }

    # Delete a message in the specified user's mailbox
    #
    # + messageId - The ID of the message 
    # + folderId - The ID of the folder where the message is in
    # + return - If success returns null otherwise the relevant error
    @display {label: "Delete Message"}
    isolated remote function deleteMessage(@display {label: "Messages ID"} string messageId,
                                           @display {label: "Folder ID"} string? folderId = ()) returns @tainted error? {
        string requestParams = folderId is string ? "/mailFolders/" + folderId + "/messages/" : "/messages/";
        requestParams += messageId;
        http:Response response = check self.httpClient->delete(requestParams);
        log:printInfo(requestParams);
        if (response.statusCode != http:STATUS_NO_CONTENT) {
            fail error("Delete operation respond with status code:" + response.statusCode.toString()); //remove
        } //check other errors body
    }

    # Send an existing draft message
    #
    # + messageId - The ID of the message 
    # + return - If success returns null otherwise the relevant error
    @display {label: "Send Draft Messages"}
    isolated remote function sendDraftMessage(@display {label: "Message ID"} string messageId) returns @tainted error? {
        string requestParams = "/messages/" + messageId + "/send";
        http:Request request = new;
        request.setHeader("Content-Length", "0");
        http:Response response = check self.httpClient->post(requestParams, request, targetType = 
        http:Response);
        log:printInfo(requestParams);
        if (response.statusCode != http:STATUS_ACCEPTED) {
            fail error("Send existing draft message operation failed with status code:" + 
            response.statusCode.toString());
        }
    }

    # Copy a message to a folder
    #
    # + messageId - The ID of the message 
    # + destinationFolderId - The ID of the destination folder  
    # + folderId - The ID of the folder where the message is in
    # + return - If success returns null otherwise the relevant error
    @display {label: "Copy Message"}
    isolated remote function copyMessage(@display {label: "Message ID"} string messageId, 
                                        @display {label: "Destination Folder ID"}string destinationFolderId, 
                                        @display {label: "Folder ID"} string? folderId = ()) returns 
                                        @tainted error? {
        string requestParams = folderId is string ? "/mailFolders/" + folderId + "/messages/" : "/messages/";
        requestParams += messageId + "/copy";
        http:Request request = new;
        request.setHeader("Content-Type", "application/json");
        request.setJsonPayload({"destinationId": destinationFolderId});
        log:printInfo(requestParams);
        http:Response response = check self.httpClient->post(requestParams, request, targetType = 
            http:Response);
        if (response.statusCode != http:STATUS_CREATED) {
            fail error("Copy message operation failed with status code:" + 
            response.statusCode.toString());
        }
    }

    # Forward a message
    #
    # + messageId - The ID of the message  
    # + comment - The comment of the forwarding message   
    # + addressList - The receivers' email list 
    # + folderId - The ID of the folder where the message is in
    # + return - If success returns the sent message detail as a message record otherwise the relevant error
    @display {label: "Forward Message"}
    isolated remote function forwardMessage(@display {label: "Message ID"} string messageId, 
                                            @display {label: "Comment"} string comment, 
                                            @display {label: "Email List"} string[] addressList, 
                                            @display {label: "Folder ID"} string? folderId = ()) 
                                            returns @tainted error? {
        string requestParams = folderId is string ? "/mailFolders/" + folderId + "/messages/" : "/messages/";
        requestParams += messageId + "/forward";
        http:Request request = new;
        request.setHeader("Content-Type", "application/json");
        ForwardParamsList parameterList = getRecipientListAsRecord(comment, addressList);
        request.setJsonPayload(parameterList.toJson());
        log:printInfo(requestParams);
        http:Response response = check self.httpClient->post(requestParams, request, targetType = 
            http:Response);
        if (response.statusCode != http:STATUS_ACCEPTED) {
            fail error("Forward message operation failed with status code:" + response.statusCode.toString());
        }
    }

    # Send a message 
    #
    # + messageContent - The message content and properties
    # + return - If success returns null otherwise the relevant error
    @display {label: "Send Message"}
    isolated remote function sendMessage(@display {label: "Message"} MessageContent messageContent) returns 
                                         @tainted error? {
        string requestParams = "/sendMail";
        http:Request request = new;
        messageContent.message.attachments = addOdataFileType(messageContent);
        request.setJsonPayload(messageContent.toJson());
        log:printInfo(messageContent.toJsonString());
        request.setHeader("Content-Type", "application/json");
        http:Response response = check self.httpClient->post(requestParams, request, targetType = 
            http:Response);
        if (response.statusCode != http:STATUS_ACCEPTED) {
            fail error("Send message operation failed with status code:" + response.statusCode.toString());
        }
    }

    # Retrieve a list of attachments
    #
    # + messageId - The ID of the message 
    # + folderId - The ID of the folder
    # + childFolderIds - The IDs of the childFolders respectively
    # + return - If success returns a ballerina stream of file attachment records otherwise the relevant error
    @display {label: "List Attachments"}
    isolated remote function listAttachment(@display {label: "Message ID"} string messageId,  
                                            @display {label: "Folder ID"}string? folderId = (), 
                                            @display {label: "Child Folder ID List"} string[]? childFolderIds = ()) 
                                            returns @tainted stream<FileAttachment, error?>|error {
        string requestParams = folderId is string ? "/mailFolders/" + folderId : "";
        requestParams += childFolderIds is () ? "" : (addChildFolderIds(childFolderIds));
        requestParams += "/messages/" + messageId + "/attachments";
        json response = check self.httpClient->get(requestParams, targetType = json);
        json[] attachmentList = let var value = response.value
            in value is json ? <json[]>value : [];
        FileAttachment[] attachments = [];
        foreach json attachment in attachmentList {
            FileAttachment fileAttachment = check attachment.cloneWithType(FileAttachment);
            attachments.push(fileAttachment);
        }
        AttachmentStream objectInstance = check new (attachments);
        stream<FileAttachment, error?> finalStream = new (objectInstance);
        return finalStream;
    }

    # Create a new mail folder in the root folder of the user's mailbox
    #
    # + displayName - The display name of the mail folder  
    # + isHidden - Indicates whether the folder should be hidden or not
    # + return - If success returns the created mail folder detail otherwise the relevant error
    @display {label: "Create Mail Folder"}
    isolated remote function createMailFolder(@display {label: "Display Name"} string displayName, 
                                              @display {label: "Is Hidden"} boolean? isHidden = ()) returns @tainted 
                                              MailFolder|error {
        string requestParams = "/mailFolders";
        http:Request request = new;
        request.addHeader("Content-Type", "application/json");
        request.setJsonPayload({displayName: displayName, isHidden: false});
        return check self.httpClient->post(requestParams, request, targetType = MailFolder);
    }

    # Create a new child mailFolder
    #
    # + parentFolderId - The ID of the parent folder
    # + displayName - The display name of the child mail folder  
    # + isHidden - Indicates whether the child folder should be hidden or not
    # + return - If success returns the created mail folder detail otherwise the relevant error
    @display {label: "Create Child Mail Folder"}
    isolated remote function createChildMailFolder(@display {label: "Parent Folder ID"} string parentFolderId, 
                                                   @display {label: "Display Name"}string displayName, 
                                                   @display {label: "Is Hidden"} boolean? isHidden = ()) 
                                            returns @tainted MailFolder|error {
        string requestParams = "/mailFolders/" + parentFolderId + "/childFolders";
        http:Request request = new;
        request.addHeader("Content-Type", "application/json");
        request.setJsonPayload({displayName: displayName, isHidden: false});
        return check self.httpClient->post(requestParams, request, targetType = MailFolder);
    }

    # Retrieve the details of a message folder
    #
    # + mailFolderId - The ID of the mail folder
    # + return - If success returns the requested mail folder details as a mail folder record otherwise the relevant 
    # error 
    @display {label: "Get Mail Folder"} 
    isolated remote function getMailFolder(@display {label: "Mail Folder ID"} string mailFolderId) returns @tainted 
                                           MailFolder|error {
        string requestParams = "/mailFolders/" + mailFolderId;
        return check self.httpClient->get(requestParams, targetType = MailFolder);
    }

    # Delete the specified mail folder or search Mail folder
    #
    # + mailFolderId - The ID of the folder by its well-known folder name, if one exists 
    # ( Eg: inbox, sentitems etc. https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0)
    # + return - If success returns null otherwise the relevant error
    @display {label: "Delete Mail Folder"}
    isolated remote function deleteMailFolder(@display {label: "Mail Folder ID"} string mailFolderId) returns @tainted 
                                              error? {
        string requestParams = "/mailFolders/" + mailFolderId;
        http:Response response = check self.httpClient->delete(requestParams, targetType = http:Response);
        if (response.statusCode != http:STATUS_NO_CONTENT) {
            fail error("Delete mail folder operation failed with status code:" + response.statusCode.toString());
        }
    }

    # Add an attachment to a message
    #
    # + messageId - The ID of the message  
    # + attachment - The File attachment detail 
    # + folderId - The ID of the folder where the message is saved in
    # + childFolderIds - The IDs of the child folders
    # + return - If success returns the added file attachment details as a record otherwise the relevant error
    @display {label: "Add File Attachment"}
    isolated remote function addFileAttachment(@display {label: "Message ID"} string messageId, 
                                               @display {label: "File Attachment"} FileAttachment attachment, 
                                               @display {label: "Folder ID"} string? folderId = (), 
                                               @display {label: "Child Folder IDs"} string[]? childFolderIds = ()) 
                                               returns @tainted FileAttachment|error {
        string requestParams = folderId is string ? "/mailFolders/" + folderId : "";
        requestParams += childFolderIds is () ? "" : (addChildFolderIds(childFolderIds));
        requestParams += "/messages/" + messageId + "/attachments";
        http:Request request = new;
        request.addHeader("Content-Type", "application/json");
        FileAttachment formattedAttachment = addOdataFileType(attachment)[0];
        request.setJsonPayload(formattedAttachment.toJson());
        return check self.httpClient->post(requestParams, request, targetType = FileAttachment);
    }

    # Get the mail folder collection directly under the root folder
    #
    # + includeHiddenFolders - Indicates whether the hidden folder should be included in the collection or not
    # + return - If success returns a ballerina stream of mail folder records otherwise the relevant error
    @display {label: "List Mail Folders"}
    isolated remote function listMailFolders(@display {label: "Include Hidden Folders"} boolean? includeHiddenFolders = 
                                                ()) 
                                             returns @tainted stream<MailFolder, error?>|error {
        string requestParams = "/mailFolders";
        requestParams += includeHiddenFolders is () ? "" : "/?includeHiddenFolders=" + includeHiddenFolders.toString();
        json response = check self.httpClient->get(requestParams, targetType = json);
        json[] mailFolderList = let var value = response.value
            in value is json ? <json[]>value : [];
        MailFolder[] mailFolders = [];
        foreach json mailFolder in mailFolderList {
            MailFolder mailFolderRecord = check mailFolder.cloneWithType(MailFolder);
            mailFolders.push(mailFolderRecord);
        }
        MailFolderStream objectInstance = check new (mailFolders);
        stream<MailFolder, error?> finalStream = new (objectInstance);
        return finalStream;
    }

    # Get the folder collection under the specified folder.
    #
    # + parentFolderId - The ID of the parent folder 
    # + includeHiddenFolders - Indicates whether the hidden folder should be included in the collection or not
    # + return - If success returns a ballerina stream of mail folder records otherwise the relevant error 
    @display {label: "List Child Mail Folders"}
    isolated remote function listChildMailFolders(@display {label: "Parent Folder ID"} string parentFolderId, 
                                                  @display {label: "Include Hidden Folder"} boolean? 
                                                  includeHiddenFolders = ()) returns  @tainted 
                                                  stream<MailFolder, error?>|error {
        string requestParams = "/mailFolders/" + parentFolderId + "/childFolders";
        requestParams += includeHiddenFolders is () ? "" : "/?includeHiddenFolders=" + includeHiddenFolders.toString();
        json response = check self.httpClient->get(requestParams, targetType = json);
        json[] mailFolderList = let var value = response.value
            in value is json ? <json[]>value : [];
        MailFolder[] mailFolders = [];
        foreach json mailFolder in mailFolderList {
            MailFolder mailFolderRecord = check mailFolder.cloneWithType(MailFolder);
            mailFolders.push(mailFolderRecord);
        }
        MailFolderStream objectInstance = check new (mailFolders);
        stream<MailFolder, error?> finalStream = new (objectInstance);
        return finalStream;
    }

    # Create a new mailSearchFolder in the specified user's mailbox
    #
    # + parentFolderId - The ID of the parent folder 
    # + mailSearchFolder - The details of the mail search folder
    # + return - If success returns the newly created mail search folder details as a MailSearchFolder record otherwise 
    #   the relevant error 
    @display {label: "Create Mail Search Folder"}
    isolated remote function createMailSearchFolder(@display {label: "Parent Folder ID"} string parentFolderId, 
                                                    @display {label: "Mail Search Folder"} MailSearchFolder 
                                                    mailSearchFolder) returns @tainted MailSearchFolder|error {
        string requestParams = "/mailFolders/" + parentFolderId + "/childFolders";
        http:Request request = new;
        request.addHeader("Content-Type", "application/json");
        json searchRequest = mailSearchFolder.toJson();
        searchRequest = check searchRequest.mergeJson({"@odata.type": "microsoft.graph.mailSearchFolder"});
        log:printInfo(searchRequest.toString());
        request.setJsonPayload(searchRequest);
        return check self.httpClient->post(requestParams, request, targetType = MailSearchFolder);
    }

    # Create an upload session that allows an the client to iteratively upload ranges of a file
    #
    # + messageId - The ID of the message
    # + attachmentName - The name of the attachment in the message
    # + file - The file path or the array of bytes of the folder or file path
    # + return - If success returns null otherwise the relevant error
    @display {label: "Add large File Attachment"}
    isolated remote function addLargeFileAttachment(@display {label: "Message ID"} string messageId, 
                                                    @display {label: "Attachment Name"} string attachmentName, 
                                                    @display {label: "Mail Search Folder"} string|byte[] file) 
                                                    returns @tainted error? {
        byte[] content = file is string ? check io:fileReadBytes("tests/sample.pdf") : file;
        AttachmentItemContent attachmentItem = {
            attachmentType: "file",
            name: attachmentName,
            size: content.length()
        };
        string requestParams = "/messages/" + messageId + "/attachments/createUploadSession";
        http:Request request = new;
        request.addHeader("Content-Type", "application/json");
        request.setJsonPayload({AttachmentItem: attachmentItem}.toJson());
        log:printInfo(requestParams.toString());
        log:printInfo(attachmentItem.toString());
        UploadSession session = check self.httpClient->post(requestParams, request, 
        targetType = UploadSession);
        log:printInfo(session.toString());
        check uploadFile(content, session);
    }

}

# Represents configuration parameters to create Azure Cosmos DB client.
#
# + clientConfig - OAuth client configuration
# + secureSocketConfig - SSH configuration
@display {label: "Connection config"}
public type Configuration record {
    http:BearerTokenConfig|http:OAuth2RefreshTokenGrantConfig clientConfig;
    http:ClientSecureSocket secureSocketConfig?;
};

isolated function uploadFile(byte[] file, UploadSession session) returns @tainted error? {
    int currentPosition = 0;
    int size = 3000000;
    boolean isFinalRequest = false;
    UploadSession uploadSession = session;
    http:Client uploadClient = check new (session?.uploadUrl.toString(), {http1Settings: {chunking: http:CHUNKING_NEVER}});
    log:printInfo(session?.uploadUrl.toString());
    while !isFinalRequest {
        http:Request request = new;
        int endPosition = currentPosition + size;
        if (endPosition > file.length()) {
            endPosition = file.length();
            isFinalRequest = true;
        }
        byte[] sliced = file.slice(currentPosition, endPosition);
        request.setBinaryPayload(sliced);
        request.addHeader("Content-Length", sliced.length().toString());
        request.addHeader("Content-Type", "application/octet-stream");
        request.addHeader("Content-Range", string `bytes ${currentPosition}-${endPosition - 1}/${file.length().toString()}`);
        log:printInfo((check request.getHeader("Content-Range")).toString());
        if (isFinalRequest) {
            http:Response response = check uploadClient->put("", request);
            if (response.statusCode != http:STATUS_CREATED) {
                fail error((check response.getJsonPayload()).toString());
            }
            break;
        }
        uploadSession = check uploadClient->put("", request, targetType = UploadSession);
        currentPosition = check 'int:fromString(uploadSession.nextExpectedRanges[0]);
    }
}

public type URIParamsList record {|
    int top?;
    string 'select?;
|};

isolated function setOptionalUriParams(URIParamsList optionalUriParameters) returns @tainted string {
    string request = "?";
    json items = optionalUriParameters.toJson();
    map<json> itemMap = <map<json>>items;
    string[] keys = itemMap.keys();
    boolean isFirst = true;
    foreach string fieldKey in keys {
        json value = itemMap.get(fieldKey);
        if (isFirst) {
            request = request + "$" + fieldKey + "=" + value.toString();
        } else {
            request = request + "&" + "$" + fieldKey + "=" + value.toString();
        }
        isFirst = false;
    }
    log:printInfo(request);
    return request;

}

isolated function addOdataFileType(MessageContent|FileAttachment messageContent) returns FileAttachment[] {
    FileAttachment[] attachments = [];
    if (messageContent is FileAttachment) {
        attachments.push(messageContent);
    } else {
        attachments = messageContent?.message?.attachments ?: [];
    }
    foreach int i in 0 ... attachments.length() {
        FileAttachment attachment = attachments.remove(0);
        FileAttachment attachmentTemp = {
            contentBytes: attachment.contentBytes,
            name: attachment.name,
            contentType: attachment.contentType,
            "@odata.type": "#microsoft.graph.fileAttachment"
        };
        if (attachment?.id is string) {
            attachmentTemp.id = attachment?.id.toString();
        }
        if (attachment?.contentId is string) {
            attachmentTemp.contentId = attachment?.contentId.toString();
        }
        if (attachment?.isInline is boolean) {
            attachmentTemp.isInline = <boolean>attachment?.isInline;
        }
        if (attachment?.lastModifiedDateTime is string) {
            attachmentTemp.lastModifiedDateTime = attachment?.lastModifiedDateTime.toString();
        }
        if (attachment?.size is int) {
            attachmentTemp.size = <int>attachment?.size;
        }
        attachments.push(attachmentTemp);
    }
    return attachments;
}

isolated function addChildFolderIds(string[] childFoldersIds) returns string {
    string requestParams = "";
    foreach string ids in childFoldersIds {
        requestParams += "/childFolders/" + ids;
    }
    return requestParams;
}
