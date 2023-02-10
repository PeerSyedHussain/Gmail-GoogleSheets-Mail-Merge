// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/**
 * Change these to match the column names you are using for email
 * recipient addresses and email sent column.
 */
const RECIPIENT_COL = "Recipient";
const EMAIL_SENT_COL = "Email Sent";

/**
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Mail Merge")
    .addItem("Send Emails (Generate new email)", "sendEmails")
    .addItem("Thread New Email (Reply to thread email)", "sendThreadEmails")
    .addToUi();
}

/**
 * Check Email sent column empty or not
 */
function checkEmailSentColumn(data) {
  const email_sent_column_empty_check = data.map((row) => {
    if (row[EMAIL_SENT_COL] === "") {
      return true;
    } else {
      return false;
    }
  });

  return email_sent_column_empty_check.includes(false);
}

/**
 * Shows alert popup
 */
function showAlert(alertMsg, subMsg) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  ui.alert("Alert!!!", `${alertMsg}` + ` ${subMsg}`, ui.ButtonSet.OK);
}

/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
 */

function getGmailTemplateFromDrafts_(subject_line) {
  try {
    // get drafts
    const drafts = GmailApp.getDrafts();
    // filter the drafts that match subject line
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    // get the message object
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    const allInlineImages = draft.getMessage().getAttachments({
      includeInlineImages: true,
      includeAttachments: false,
    });
    const attachments = draft
      .getMessage()
      .getAttachments({ includeInlineImages: false });
    const htmlBody = msg.getBody();

    // Creates an inline image object with the image name as key
    // (can't rely on image index as array based on insert order)
    const img_obj = allInlineImages.reduce(
      (obj, i) => ((obj[i.getName()] = i), obj),
      {}
    );

    //Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^>]+>', "g");
    const matches = [...htmlBody.matchAll(imgexp)];

    //Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach((match) => (inlineImagesObj[match[1]] = img_obj[match[2]]));

    return {
      message: {
        subject: subject_line,
        text: msg.getPlainBody(),
        html: htmlBody,
      },
      attachments: attachments,
      inlineImages: inlineImagesObj,
    };
  } catch (e) {
    showAlert("Oops - can't find Gmail draft", "");
    return;
  }

  /**
   * Filter draft objects with the matching subject linemessage by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} GmailDraft object
   */
  function subjectFilter_(subject_line) {
    return function (element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    };
  }
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
 */
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, (key) => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
 */
function escapeData_(str) {
  return str
    .replace(/[\\]/g, "\\\\")
    .replace(/[\"]/g, '\\"')
    .replace(/[\/]/g, "\\/")
    .replace(/[\b]/g, "\\b")
    .replace(/[\f]/g, "\\f")
    .replace(/[\n]/g, "\\n")
    .replace(/[\r]/g, "\\r")
    .replace(/[\t]/g, "\\t");
}

/**
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 */
function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  // option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine) {
    subjectLine = Browser.inputBox(
      "Mail Merge",
      "Type or copy/paste the subject line of the Gmail " +
        "draft message you would like to mail merge with:",
      Browser.Buttons.OK_CANCEL
    );

    if (subjectLine === "cancel" || subjectLine == "") {
      // If no subject line, finishes up
      return;
    }
  }

  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift();

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {})
  );

  // check email sent column empty or not
  const showAlertForEmailColumnNotEmpty = checkEmailSentColumn(obj);

  if (showAlertForEmailColumnNotEmpty) {
    showAlert("Clear email sent column & retry it!!!", "");
  } else {
    // Creates an array to record sent emails
    const out = [];

    // Loops through all the rows of data
    obj.forEach(function (row, rowIdx) {
      // Only sends emails if email_sent cell is blank and not hidden by a filter
      if (row[EMAIL_SENT_COL] == "") {
        try {
          const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

          // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
          // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
          // Uncomment advanced parameters as needed (see docs for limitations)
          GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
            htmlBody: msgObj.html,
            // bcc: 'a.bbc@email.com',
            // cc: 'a.cc@email.com',
            // from: 'an.alias@email.com',
            // name: 'name of the sender',
            // replyTo: row[RECIPIENT_COL],
            // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages,
          });
          // Edits cell to record email sent date
          out.push([new Date()]);
        } catch (e) {
          // modify cell to record error
          out.push([e.message]);
        }
      } else {
        out.push([row[EMAIL_SENT_COL]]);
      }
    });

    // Updates the sheet with new data
    sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
  }
}

/**
 * Thread new email
 */
function sendThreadEmails(
  subjectLine,
  sheet = SpreadsheetApp.getActiveSheet()
) {
  // option to skip browser prompt if you want to use this code in other projects

  if (!subjectLine) {
    subjectLine = Browser.inputBox(
      "Thread Loop Mail Merge",
      "Type or copy/paste the subject line of the Gmail " +
        "draft message you would like to mail merge with:",
      Browser.Buttons.OK_CANCEL
    );

    if (subjectLine === "cancel" || subjectLine == "") {
      // If no subject line, finishes up
      return;
    }
  }
  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();

  // Fetches displayed values for each row in the Range HT Andrew Roberts
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift();

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {})
  );

  // check email sent column empty or not
  const showAlertForEmailColumnNotEmpty = checkEmailSentColumn(obj);

  if (showAlertForEmailColumnNotEmpty) {
    showAlert("Clear email sent column & retry it!!!", "");
  } else {
    // Creates an array to record sent emails
    const out = [];

    const special_character_error_row = [];

    // Loops through all the rows of data
    obj.forEach(function (row, rowIdx) {
      // Only sends emails if email_sent cell is blank and not hidden by a filter
      if (row[EMAIL_SENT_COL] == "") {
        try {
          const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

          var filteringWithOriginalEmailSubject =
            row["Original Email Subject For Filtering"];

          if (filteringWithOriginalEmailSubject !== "") {
            // var placeholders = filteringWithOriginalEmailSubject.match(/\$(.*?)\$/g)    // matching $ in text
            var placeholders;

            if (filteringWithOriginalEmailSubject.includes("{{")) {
              placeholders =
                filteringWithOriginalEmailSubject.match(/\{(.*?)\}}/g); // matching {} in text

              if (placeholders.length > 0) {
                placeholders.forEach((placeholder) => {
                  //Placeholder - {{First Name}}
                  var phText = placeholder.substring(2, placeholder.length - 2);
                  //phText = First Name
                  if (row[phText]) {
                    filteringWithOriginalEmailSubject =
                      filteringWithOriginalEmailSubject.replace(
                        placeholder,
                        row[phText]
                      );
                  }
                });
              } else {
                showAlert(
                  "Check Original Email Subject For Filtering With Placeholders!!",
                  ""
                );
              }
            } else {
              showAlert(
                "Check Original Email Subject For Filtering With Placeholders!!",
                ""
              );
            }
            /**
             * Flow of email sending based on search query
             */
            var unanswered = GmailApp.search(
              `'in:sent to:(${row.Recipient}) subject:"${filteringWithOriginalEmailSubject}"'`
            );

            if (unanswered.length > 0) {
              var messages = unanswered[0].getMessages();

              var lastMsg = messages[messages.length - 1]; // get the last message in the thread
              // set your followup text.
              // + Note that message.reply() & message.forward()` don't append the thread to the reply, so you'll have to do this yourself -- both text & HTML
              var followupText =
                "" + lastMsg.getPlainBody().replace(/^/gm, "> ");
              var followupHTML =
                `<div class="gmail_quote"> ${msgObj.html}` +
                lastMsg.getBody() +
                "</div>";
              var email_re = new RegExp(
                Session.getActiveUser().getEmail(),
                "i"
              );

              if (email_re.test(lastMsg.getFrom())) {
                lastMsg.forward(lastMsg.getTo(), {
                  subject: lastMsg.getSubject(),
                  htmlBody: followupHTML,
                });
              } else {
                lastMsg.reply(followupText, { htmlBody: followupHTML });
              }
              out.push([new Date()]); //actual sent date filled in specific index
            } else {
              var obj = {};
              obj.row = row;
              obj.rowId = rowIdx + 2; // add 2 for matching the sheet row number

              special_character_error_row.push(obj);

              out.push([""]); //just make empty array filled in specific index which are not able to send email
            }
          } else {
            showAlert("Enter Original Email Subject For Filtering", "");
          }
        } catch (e) {
          // modify cell to record error
          out.push([e.message]);
        }
      } else {
        out.push([row[EMAIL_SENT_COL]]);
      }
    });

    if (special_character_error_row.length > 0) {
      showAlert(
        "Remove special characters like [|, /, ^] and retry it.",
        `Fix Issues in First name: ${special_character_error_row.map(
          (each) => each.row["First name"]
        )} Or Fix Issues in row number: ${special_character_error_row.map(
          (each) => each.rowId
        )}`
      );
    }

    // Updates the sheet with new data
    sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
  }
}