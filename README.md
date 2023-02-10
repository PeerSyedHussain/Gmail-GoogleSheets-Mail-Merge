What is Google Apps Script?

Google Apps Script home page defines itself as follows.

A cloud-based JavaScript platform that lets you integrate with and automate tasks across Google products.

Google has done an excellent job of hiding all the complexities of connecting to different Google APIs. An Apps Script user has to call a couple of pre-defined functions and viola, you can do miracles!!!

Not impressed yet? Let me give some sample use cases of Apps Script.

    Create calendar events based on event details on a spreadsheet and send registration forms to the invitees of the events.
    Create a pretty invoice and email it using the data in a spreadsheet.
    Create a Google calendar event using Google Chat.

Now that we have a good understanding of what Apps Script is let’s see how we can utilize Apps Script to achieve our target.

Create a mail merge with Gmail & Google Sheets:

About this solution:

Automatically populate an email template with data from Google Sheets. The emails are sent from your Gmail account so that you can respond to recipient replies.

Important: This mail merge sample is subject to the email limits described in Quotas for Google services.

How it works:

You create a Gmail draft template with placeholders that correspond to data in a Sheets spreadsheet. Each column header in a sheet represents a placeholder tag. The script sends the information for each placeholder from the spreadsheet to the location of the corresponding placeholder tag in your email draft.

Apps Script services:

This solution uses the following services:

    Gmail service–Gets, reads, and sends the draft email you want to send to your recipients.
        If your email includes unicode characters like emojis, use the Mail service instead. Learn how to update the code to include unicode characters in your email.
    Spreadsheet service–Fills in the email placeholders with the personalized information for each of the recipients.

For more info: https://developers.google.com/apps-script/samples/automations/mail-merge

https://developers.google.com/static/apps-script/samples/images/mail-merge.gif