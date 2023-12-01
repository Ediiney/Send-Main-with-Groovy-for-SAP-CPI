/** This script fetches an access token from the OAuth2 Authorization Code credential whose name is provided
 * by the property "OAUTH2_AUTHORIZATION_CODE_CRED", sets the access token as Bearer token in the header "Authorization",
 * and creates a body in JSON format which can be sent to the Microsoft Graph API "https://graph.microsoft.com/v1.0/me/sendMail" 
 * via the HTTP receiver adapter.
 * 
 * The incoming message body is used as mail body and it is assumed the incoming message body has a charset encoding of UTF-8.
 * 
 * The script is parametrized by the following properties:
 * - "OAUTH2_AUTHORIZATION_CODE_CRED": name of the OAuth2 Authorization Code credential created before (in our example: "SEND_MAIL_VIA_GRAPH_API")
 * - "MAIL_SAVE_TO_SENT_ITEMS": indicator whether to save the sent mail into the "Sent Items" folder of the mail account
 * - "MAIL_RECIPIENT": e-mail address of the mail recipient
 * - "MAIL_CC_RECIPIENT": optional e-mail address of the cc recipient
 * - "MAIL_SUBJECT": subject of the mail
 * - "MAIL_ATTACHMENT": optional property containing the base 64 encoded mail attachment
 * - "MAIL_ATTACHMENT_CONTENT_TYPE": content type of the mail attachment (example: "text/plain"), only necessary if MAIL_ATTACHMENT is provided
 * - "MAIL_ATTACHMENT_NAME": name of the mail attachment, only necessary if MAIL_ATTACHMENT is provided
 * 
 * See also API description from Microsoft: https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
 */

/* Refer the link below to learn more about the use cases of script.
https://help.sap.com/viewer/368c481cd6954bdfa5d0435479fd4eaf/Cloud/en-US/148851bf8192412cba1f9d2c17f4bd25.html

If you want to know more about the SCRIPT APIs, refer the link below
https://help.sap.com/doc/a56f52e1a58e4e2bac7f7adbf45b2e26/Cloud/en-US/index.html */
import com.sap.gateway.ip.core.customdev.util.Message
import java.util.HashMap
import com.sap.gateway.ip.core.customdev.util.Message
import com.sap.it.api.securestore.SecureStoreService
import com.sap.it.api.securestore.AccessTokenAndUser
import com.sap.it.api.securestore.exception.SecureStoreException
import com.sap.it.api.ITApiFactory
import java.nio.charset.StandardCharsets
import groovy.json.JsonOutput
import java.util.Base64
import java.util.LinkedHashMap

def Message processData(Message message) {
     def properties = message.getProperties();
     
    SecureStoreService secureStoreService = ITApiFactory.getService(SecureStoreService.class, null);
    // fetch the name of the OAuth2 Authorization Code credential from the property "OAUTH2_AUTHORIZATION_CODE_CRED"
    String  oauth2AuthorizationCodeCred = properties.get("OAUTH2_AUTHORIZATION_CODE_CRED");
    if (oauth2AuthorizationCodeCred == null || oauth2AuthorizationCodeCred.isEmpty()){
        throw new IllegalStateException("Property OAUTH2_AUTHORIZATION_CODE_CRED is not specified")
    }
    // fetch access token 
    AccessTokenAndUser accessTokenAndUser = secureStoreService.getAccesTokenForOauth2AuthorizationCodeCredential(oauth2AuthorizationCodeCred);
    String token = accessTokenAndUser.getAccessToken();
    
    // set the Authorization header with the access token
    message.setHeader("Authorization", "Bearer "+token);
    message.setHeader("Content-Type","application/json");
    message.setHeader("MIME","text/plain");
    
    MailMessage mailMessage = new MailMessage()
     String  mailSubject = properties.get("MAIL_SUBJECT");
    if (mailSubject == null || mailSubject.isEmpty()){
        throw new IllegalStateException("Property MAIL_SUBJECT is not specified")
    }
    mailMessage.subject = mailSubject
    
    // We assume here that the mail content is provided in the message body as byte array with UTF8 encoding
    // and that the content type is provided in the property MAIL_BODY_CONTENT_TYPE
    //     Possible values vor the content type are  "text" and "html".
    String  contentType = properties.get("MAIL_CONTENT_TYPE");
    if (contentType == null || contentType.isEmpty()){
        throw new IllegalStateException("Property MAIL_CONTENT_TYPE is not specified")
    }
    byte[] messageBody =  message.getBody(byte[].class)
    if (messageBody == null){
         throw new IllegalStateException("Message body is null")
    }
    MailBody mailBody = new MailBody()
    mailBody.contentType = contentType
    mailBody.content = new String(messageBody,StandardCharsets.UTF_8) // we assume here that the message body is encoded in UTF8
    
    mailMessage.body = mailBody
    
    String  mailRecipient = properties.get("MAIL_RECIPIENT");
    if (mailRecipient == null || mailRecipient.isEmpty()){
        throw new IllegalStateException("Property MAIL_RECIPIENT is not specified")
    }
    EmailAddress emailAddress = new EmailAddress()
    emailAddress.address = mailRecipient
    
    MailRecipient recipient = new MailRecipient()
    recipient.emailAddress = emailAddress
    
    // you can also add more than one recipient
    mailMessage.toRecipients.add(recipient)
   
      //mailCcRecipient is optional
    String  mailCcRecipient = properties.get("MAIL_CC_RECIPIENT");
    if (mailCcRecipient != null && (! mailCcRecipient.isEmpty())){
        EmailAddress ccEmailAddress = new EmailAddress()
        ccEmailAddress.address = mailCcRecipient
        
        MailRecipient ccRecipient = new MailRecipient()
        ccRecipient.emailAddress = ccEmailAddress
        
        mailMessage.ccRecipients.add(ccRecipient)
    }
    
    // attachment is optional, attachment must be base 64 encoded
    String  attachment = properties.get("MAIL_ATTACHMENT");
    if (attachment != null && (! attachment.isEmpty())){
        String  attachmentContentType = properties.get("MAIL_ATTACHMENT_CONTENT_TYPE");
        if (attachmentContentType == null || attachmentContentType.isEmpty()){
            throw new IllegalStateException("Property MAIL_ATTACHMENT_CONTENT_TYPE is not specified")
        }
        String  attachmentName = properties.get("MAIL_ATTACHMENT_NAME");
        if (attachmentName == null || attachmentName.isEmpty()){
            throw new IllegalStateException("Property MAIL_ATTACHMENT_NAME is not specified")
        }
        Map<String,String> mailAttachment = new LinkedHashMap<>()
        mailAttachment.put("@odata.type","#microsoft.graph.fileAttachment")
        mailAttachment.put("name", attachmentName)
        mailAttachment.put("contentType", attachmentContentType)
        mailAttachment.put("contentBytes", attachment )
        // you can also add more than one attachment
        mailMessage.attachments.add(mailAttachment)
    }
    
    Mail mail = new Mail()
    mail.message = mailMessage;
    String  saveToSentItems = properties.get("MAIL_SAVE_TO_SENT_ITEMS");
    if (saveToSentItems == null || saveToSentItems.isEmpty()){
        // default value is true
        mail.saveToSentItems = true
    } else {
        boolean saveToSentItemsBoolean = Boolean.valueOf(saveToSentItems);
        mail.saveToSentItems = saveToSentItemsBoolean
    }
   
    String jsonBody = JsonOutput.toJson(mail)
    message.setBody(jsonBody.getBytes(StandardCharsets.UTF_8));
    
    return message;
}

class MailBody{
    
    String contentType = "text/html"
    String content
}

class EmailAddress{
    String address
}

class MailRecipient{
    
    EmailAddress emailAddress
    
}


class MailMessage{
    
    String subject
    MailBody body
    List<MailRecipient> toRecipients = new LinkedList<MailRecipient>()
    List<MailRecipient> ccRecipients = new LinkedList<MailRecipient>()
    List<Map<String,String>> attachments = new LinkedList<Map<String,String>>()
    /* If you want to use "internetMessageHeaders" as described in https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http#example-2-create-a-message-with-custom-internet-message-headers-and-send-the-message, then you enhance the MailMessage class. */
}

class Mail {
    
    MailMessage message
    boolean saveToSentItems
}