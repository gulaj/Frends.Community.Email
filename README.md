# Frends.Community.Email
Frends task for sending emails

- [Installing](#installing)
- [Tasks](#tasks)
  - [Send Email](#sendemail)
- [License](#license)
- [Building](#building)
- [Contributing](#contributing)
- [Change Log](#change-log)

# Installing
You can install the task via FRENDS UI Task View or you can find the nuget package from the following nuget feed
'Nuget feed coming at later date'

Tasks
=====

## SendEmail

### Task Parameters

### Message
Settings for message properties

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| To | string | Recipient addresses separated by , or ; | jane.doe@somedomain.com |
| From | string | The email address the message is sent from | john.doe@somedomain.com |
| Sender Name | string | Name of the sender | Frends errors |
| Subject | string | Subject of the message | Hello Jane |
| Message | string | Body of the message | You've got mail! |
| Is message html | bool |  Indicates whether the mail message body is in Html | false |
| Message encoding| string | Encoding of message body and subject. Use following table's name column for other options. https://msdn.microsoft.com/en-us/library/system.text.encoding(v=vs.110).aspx#Anchor_5 | utf-8 |


### Attachments

Settings for included attachments

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Attachment type | enum { FileAttachment, AttachmentFromString } | Chooses if the attachment file is created from a string or copied from disk  | FileAttachment |
| File content | string |  Usable with Attachment type AttachmentFromString. Sets the content of the attachment file | Lorem ipsum... |
| File name | string | Usable with Attachment type AttachmentFromString. Sets the name of the attachment file | error.txt |
| File path | string | Usable with Attachment type FileAttachment. Attachment file's path. Uses Directory.GetFiles(string, string) as a pattern matching technique. See https://msdn.microsoft.com/en-us/library/wz42302f(v=vs.110).aspx. Exception: If the path ends in a directory, all files in that folder are added as attachments. | C:\\temp\\*.csv |
| Throw exception if attachment not found | bool | Usable with Attachment type FileAttachment. If set true and no files match the given path, an exception is thrown | true |
| Send if no attachments found | bool | Usable with Attachment type FileAttachment. If set true and no files match the given path, email will be sent nevertheless | false |

### SMTP Settings

Settings for connecting to SMTP server

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Smtp server | string | SMTP server address | smtp.somedomain.com |
| Port | int | SMTP server port | 25 |
| Use ssl| bool | Set this true if SMTP expects to be connected using SSL | false |
| Use windows authentication | bool | Set this true if you want to use windows authentication to authenticate to SMTP server | false |
| User name| string | Usable when windows authentication is set false. Use this username to log in to the SMTP server | user |
| Password | string | Usable when windows authentication is set false. Use this password to log in to the SMTP server | password |


### Result
| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| EmailSent | bool | Returns true if email message has been sent | true |
| StatusString| string | Contains information about the task's result. | No attachments found matching path \"C:\\temp\\*.csv\". No email sent. |


# License

This project is licensed under the MIT License - see the LICENSE file for details

# Building

Clone a copy of the repo

`git clone https://github.com/CommunityHiQ/Frends.Community.Email.git`

Restore dependencies

`nuget restore frends.community.email`

Rebuild the project

Run Tests with nunit3. Tests can be found under

`Frends.Community.Email.Tests\bin\Release\Frends.Community.Email.Tests.dll`

Create a nuget package

`nuget pack nuspec/Frends.Community.Email.nuspec`

# Contributing
When contributing to this repository, please first discuss the change you wish to make via issue, email, or any other method with the owners of this repository before making a change.

1. Fork the repo on GitHub
2. Clone the project to your own machine
3. Commit changes to your own branch
4. Push your work back up to your fork
5. Submit a Pull request so that we can review your changes

NOTE: Be sure to merge the latest from "upstream" before making a pull request!

# Change Log

| Version             | Changes                 |
| ---------------------| ---------------------|
| 1.0.0 | Initial version of SendEmail |
