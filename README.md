# ITG Scripts

This repo includes a collection of scripts for ITG that either link into the ITG API to modify data in our documentation, or helper scripts to help verify ITG data. This does not include auto documentation scripts. Some larger scripts also have their own repo.

## Scripts
The following scripts are included:

### AD User Compare
This script was a precursor to the User Audit and is largely not needed anymore.

This scripts isn't entirely ITG related. It checks a list of users in a csv against AD (and optionally O365) to find which users on your list don't exist in AD. You can export a list of users from ITG or take an onboarding list and compare the user list against AD. Your CSV should have the following fields: FullName, Username, and Email. Not all fields must be filled in but the headers must exist. You can leave the column empty if you wish. You must run this script on that customers AD server.

When ran, the script will let you choose the CSV and then it will check each user against various AD properties to find matches. When complete it will return a list of users that were not found in AD. 

### Bulk Edit - Configurations
This script has limited use for us as you can only edit a small subset of data if a configuration is synced with RMM. As most of our devices are synced, this can only be used on a handful of devices.

The script can be used to mass edit a field on configurations. By default it will change all of the configurations but you can easily add a filter to the `Get-ITGlueConfigurations` line to only select a subset of configurations. You can update the `$UpdatedConfig` variable to change what field(s) gets modified. Before use, you will need to fill in an ITG API key and the ITG organization ID for the customer you are editing. 

### Contact Cleanup
This script cleans up contacts in Autotask based on data in IT Glue. Currently it runs through each phone number and formats them using Google's `libphonenumber` library. Additionally, it cleans up title fields and deactivate's any Autotask contacts that are Terminated in IT Glue.

Prerequisites: This code uses some inline C# code for the phone number formatting. You will need the .NET SDK to run this. Find that here: https://dotnet.microsoft.com/download
Additionally, you will need to open the PhoneNumberFormatter project and then have it install all dependencies. This will install the `libphonenumber` library. Open the PhoneNumberFormatter folder, then in a terminal window run `dotnet build`; this will install all dependencies. Next run `dotnet nuget locals global-packages -l` to find the location of the NuGet packages, navigate to that location. Find the `libphonenumber-csharp` folder and then navigate into the highest version. Navigate into the `lib` then `net46` folders. Get the path to the `PhoneNumbers.dll` file, use this path for the `$phoneNumbersDLLPath` constant.

### ITG Flexible Asset Fields Backup
This script backs up all flexible assets in ITG. It will create a folder with a date/timestamp each time it runs a backup, and in the folder is a separate json file for each flexible asset. These json files contain info on the flexible asset itself, and all fields within it. It should be all the info needed for recreating these flexible asset templates. Once complete, it will cleanup old backups, keeping a configurable amount of days worth of backups. The script is designed to be ran once a day, but you can run it more or less often.

### Quickpass Cleanup
This script is used to cleanup passwords in Quickpass and to setup default settings. It is intended for a Quickpass setup that is connected to ITG and Autotask. It is best used in conjunction with the Password Cleanup script below, which will cleanup data in ITG, helping this script with fixing password matching to ITG and Autotask. This cleanup focuses on fixing matches between QP and ITG/Autotask, phone numbers and emails, and setting up default settings. If any issues cannot be fixed manually, an email will be sent with a list of those issues for manual fixing.

These are the main things the cleanup will attempt to fix:
- Updates password matches between Quickpass and ITGlue/Autotask for any unmatched passwords
- Alerts on any companies that are not matched with ITG/Autotask
- Updates phone numbers in QP from Autotask contacts (will overwrite from ITG the first time unless the user answered an onboarding email)
  - If not updating from Autotask, it will clean up the phone number if it doesn't properly start with the country code
- Updates emails from Autotask contacts if missing in QP
- Sets up default rotation settings on any organizations without this setup
- Finds any AD servers (from the ITG Active Directory flexible asset) without a QP agent and reports these for install

To setup the script up you will need to fill in an ITG API key (with full password access) and optionally, an email API key from the Automation Mailer (or SendGrid or something similar). Setup the EmailFrom and EmailTo if setting up the email API key. You will also need to create a non-SSO account in Quickpass, it can have MFA. If SSO is setup then to bypass SSO, this account must have the Super or Owner login role. If the QP account has MFA, add the MFA secret and be sure to include the `GoogleAuthenticator.psm1` file in the same folder as this script. The `$EmailCategories` are the ITG password categories that are associated with email account passwords. You will also need to set the name of your Active Directory flexible asset in ITG, it will look at the `ad-servers` field to find any AD servers that are missing the QP agent.

### Password Cleanup
This script can be used to cleanup password documentation in ITG. There are 2 versions, one manual version that allows you to step through and manually approve various changes (others are done automatically), and an automatic version that can be run daily unattended to fix most issues that don't need manual intervention. The automatic version will send out an email with any manual changes that need to be looked into or completed. This cleanup has mainly been created to help cleanup naming/categorization and to assist in matching with Quickpass, and cleaning up issues that Quickpass creates in our documentation. This is focused on user passwords and requires an AD and Email asset to be setup in ITG. Additionally, we use a password naming convention of `Type - Asset or User` (e.g. "AD & O365 - Test User"), this script heavily relies on that naming convention and will set password names to use that convention.

These are the main things the cleanup will attempt to fix:
- The passwords category and the type of password in the naming (based on our naming convention, e.g. "AD - Test User" or "AD & O365 - Test User")
  - For this the cleanup queries the AD and Email assets in ITG to determine what is setup and if AD/O365 sync is setup
- Username standardization to help with matching in Quickpass
- Fixes naming and categorization of Quickpass created passwords (it then renames the "Created by Quickpass" note to "Created by Quickpass - Cleaned")
- Looks for duplicate passwords created by Quickpass, updates the existing password the new username if applicable, deletes the new password created by Quickpass, and suggests you manually update the Quickpass match to the correct password
  - You can setup Quickpass credentials in the configuration of the Automatic version, if so, the script will try to automatically update the Quickpass to ITG match after deleting the duplicate
- Properly labels and categorizations O365 Global Admin passwords
- Automatic only: Attempts to match passwords (as related items) to contacts and configurations. This saves the last time this was done and only looks at new passwords since the last run as it can take quite a long time to run

To setup the script up you will need to fill in an ITG API key (with full password access) and an email API key from the Automation Mailer (or SendGrid or something similar). Setup the EmailFrom and EmailTo if setting up the Automatic script. In the Automatic script, you can optionally setup the Quickpass account integration. To do so, create a non-SSO account in Quickpass, it can have MFA. If SSO is setup then to bypass SSO, this account must have the Super or Owner login role. If the QP account has MFA, add the MFA secret and be sure to include the `GoogleAuthenticator.psm1` file in the same folder as this script. The `$PasswordCategoryIDs` are the ITG category ID's that you want associated with different types of passwords. The `OldEmail` and `OldEmailAdmin` we used as we wanted to move O365 Global Admin passwords to a new category, but in most cases these can be ignored and that code can be removed. You will also need to set the names of your AD and Email flexible assets in ITG. The `$PasswordCatTypes_ForMatching` and `$ContactTypes_` variables can be modified to match the category types you use in ITG.

### Password Converter
When using MyGlue we have found that you cannot restrict security access to embedded passwords. Embedded passwords are tied to the security of the asset they are embedded in. This does not provide the control we require so we decided to primarily use general passwords instead. This script can be used to mass convert embedded passwords into general passwords.

There is no way I could find natively in ITG to convert an embedded password to a general passwords so this script create a brand new passwords, adds related items, and then deletes the old password. If the original embedded password was linked directly from another asset's form field (e.g. a form field on a wifi asset holds this password), the script cannot update that and will instead at the end give a list of manual fixes that are required. While it would be possible to have the script modify these other assets, it would be rather tedious as you'd have to handle every single different asset as well as flexible assets. Flexible assets would be the most problematic as if you don't include all the existing fields, data will be lost. When it lists the manual changes that are required, be careful to get these fixed correctly as once you close this window, the data will be lost permanently. 

To setup the script up you will need to fill in an ITG API key (with full password access) and the ITG organization ID of the customer you want to convert passwords for. You can also modify the limit; this is how many passwords it will convert in one batch. Manual fixes will be easier if you don't make this amount too large.

There is an additional script called "Password Convert - All Contacts" that will run for all contacts across all organizations if you need a quick way to target all of your documentation.

:heavy_exclamation_mark::heavy_exclamation_mark::heavy_exclamation_mark: If you are unsure of how this tool works, I would suggest testing it first on a dummy organization with some fake embedded passwords. If not used correctly, important password data can be lost!

### Password Import Tool
This tool can be used to mass import a list of passwords from a Confluence HTML page. It parses the HTML page for a table although it would be easy to modify it for any csv or table of passwords. To use it, simply save the existing webpage to a file, run this tool, and then import the HTML page. The script will walk you through a number of import options. 

**How it Works**
1. It will start by having you choose the import type. The script currently supports options for Contacts and Configurations. The option you choose on this form will determine naming of the resulting password, if it is embedded or general, and what it gets linked to by default. More types can be added by modifying the `$ImportTypes` variable. 
2. After choosing a type, the script will have you choose the HTML page. It will parse this for any tabular data, then will give you the option to select which table to import from. If it can't find a name, it will call the table "No Name #X". An example of the table is shown to make it easier to find the correct table. Be sure to choose the table you want correctly! 
3. The script will then parse the table giving you options to choose the username/naming/password/embedded asset name columns. It will then filter out the required fields and attempt to match the related contact or configuration. If it cannot match the contact/configuration, it will give you the option to manually choose this. 
4. Finally it will check if the password already exists and if not, it will upload each password it found. If the password already exists, it will output a warning or overwrite it if `$OverwritePasswords = $true`. 

To setup the script up you will need to fill in an ITG API key (with full password access), the ITG organization ID of the customer you want to convert passwords for, and the device prefix (e.g. STS). You can optionally add a folder ID to place all of the passwords in a specific passwords folder. Get this from the ITG url of that folder. Additionally, set `$OverwritePasswords` to `$true` if you want to overwrite existing passwords. `$ImportTypes` and `$PresetUsernames` can be used to add more password types. 

This script relies on multiple xaml forms built in Visual Studio. These are found in the Forms folder and if the script cannot find these forms, it will not work. Full solution files for Visual Studio are included so that the forms can be modified easily. 

### Notes Import Tool
This tool can be used to mass import notes on contacts or configurations from a Confluence HTML page. It is a modified version of the above Password Import Tool and requires the same forms. You may notice that some of the forms directly reference passwords, just ignore that. This script works almost the exact same as the Password Import Tool, you just select the column with the notes and the columns for matching, then it will match assets (either contacts or configurations) and upload the new notes.