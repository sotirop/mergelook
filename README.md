# mergelook

Using **mergelook** you can send multiple emails, like using the mail merge feature of MS Word, but with attachments.

## More info

There are times when you want to send more or less the same -based on a template- to multiple recipients. MS Word in combination with MS Excel and MS Outlook offers this functionality. However, the functionality is limited to sending emails without attachments. Using **mergelook** you can send emails with attachments to multiple recipients.

## Instructions

In order to use **mergelook**, you need to have MS Office installed. At least MS Excel and MS Outlook are needed. From this repository, you need the following two files:

+ `mergelook.xlsm`
+ `message.oft`

`message.oft` is the template of the email to be sent to multiple recipients.
Special words in the template will be replaced with information in the `mergelook.xlsm` file.

## Attention
The code should not be used for production use. Before using it, put Outlook in Offline Mode using the following instructions:

1. Open Outook
2. Go to "Send / Receive" tab
3. Enable "Work Offline" mode. This button should be pressed
![Work Offline][wo]

[wo]: https://raw.githubusercontent.com/sotirop/mergelook/master/Work-Offline.png "Work Offline"

When the information in the `mergelook.xlsm` is complete, press the "Send Emails" button there and watch the Outlook Outbox folder filling in with messages. Open some messages to test if everything went OK. If you are happy, press the "Send All" in Outlook so all messages can leave for their destination.