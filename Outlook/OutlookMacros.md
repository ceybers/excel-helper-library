# Outlook Macros
## `EmailsFromMeeting.bas`
- Returns a list of all the recipients (name and email address) of the selected meeting.
- Outputs to VBA IDE Immediate window.
- A meeting must be open in an active Inspector window.

## `LANDateStamp.bas`
- Returns a stamp for the selected email containing the LAN ID (or email address for external users) and the date/time.
- An email must be open in an active Inspector window, or exactly one email must be selected in an active Explorer window.
  
## `RFC2822.bas`
- `GetMessageIDofEmail` returns the Message ID of an email.
  - The active window will be used if it is an email Inspector.
  - The selected email will be used if the active Explorer has exactly one email selected. 
- `ShowEmailByMessageID` displays the email with the given Message ID.
  - The email will be searched for in the currently selected folder and its subfolders.

## `SaveAttachmentsToFolders.bas`
- Saves all the attachments in all the selected emails into a folder.
- The folder is created in the user's temporary folder and is automatically opened in File Explorer.
- A subfolder is created for each email where attachments are saved.
- Embedded or attached images are not saved (including images in signatures).

## Notes
### Uniquely Identifying Emails
- `PR_INTERNET_MESSAGE_ID` is an option.
  - `http://schemas.microsoft.com/mapi/proptag/0x1035001F`
  - Possibly consistent across all mailboxes?
  - Is only set once a message is saved (or sent) while in Online mode.
    - i.e., not available on blank emails and unsaved drafts, or saved drafts while offline.
- `PR_SEARCH_KEY ` might also be an option.

## External References
- [Items.Restrict method (Outlook) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.restrict)
- [How to uniquely identify Mail Objects in Outlook using VBA across multiple email-IDs - Stack Overflow](https://stackoverflow.com/q/70846809/)
- [c# - Message Id (PR_INTERNET_MESSAGE_ID_W_TAG) Not Available in Event Handler After Sending An Email - Stack Overflow](https://stackoverflow.com/q/54432156/)