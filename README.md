# SIMS-Academic-Progress-Report-Sender
A Windows tool for sending SIMS academic progress reports in bulk directly to parent emails via any domain joined mailbox. Authenticates via MSAL and sends personalised emails with attached reports through the Microsoft Graph API.

This exists because bulk report distribution isn't a standard feature in SIMS without a significant price tag. I was asked to put it together to save the office from manually mailing each report individually. 

**Features:**
- Test dry run with a log of where emails would be sent
- Filename based matching system with duplicate email detection, one per unique email address
- Customisation of email subject and body
- Execution log
- Send log local to reports folder to prevent duplicate emails
- App config to save details

**Setup**

Setup is straightforward but you will need a domain admin and someone with SIMS report access. After the initial setup this can be used by anyone with SIMS reports access.

You'll need to access the Microsoft Entra Admin Panel and register a new app. Note the **Application (Client) ID** and **Directory (Tenant) ID**.

Next, give the application the **Mail.Send.Shared delegated permission** and grant admin consent.

On Microsoft 365 Admin Center, give the user (e.g. user1@domainemail.com) **Send As** permissions for the mailbox you'd like to send from (e.g. sharedmailbox1@domainemail.com).

**Note that if your school uses a service such as LGFL for your shared mailboxes you will need to grant the user permissions in their web portal also.**

You will then need to create a report on SIMS containing the following columns in this exact order:
Student Firstname, Student Lastname, Primary Email, Home Email, Work Email

**The column order is important, names must come first. The three email columns can contain any email type as long as they are the last three columns.**

**Running the Tool**

Run the report you've created in SIMS and export it as a CSV.

Create a folder for your academic report files and place them inside.

Within the tool, enter the Tenant ID and Client ID you saved earlier, select your CSV and reports folder, enter the send from email address (e.g. sharedmailbox1@domainemail.com) and adjust the email subject and body if needed.

It is recommended to replace all parent email addresses in your CSV with your own for an initial test using a small batch of around 5 reports.

Run the tool with **Enable Dry Run** checked first. It will show you which files it has processed, whether a match was found in the CSV, and where the email would have been sent and from which account.

When you are satisfied, uncheck dry run and run the tool. It will ask you to login to your Microsoft account. Use the user account with permissions 'user1@domainemail.com'. If successful, **it will send the emails**. As soon as at least one email sends successfully the send log will be updated. Running the tool again will not resend emails already in the log. You can continue adding reports to the same folder and re-running the tool, clearing the sent log when the previous reporting cycle has ended.
