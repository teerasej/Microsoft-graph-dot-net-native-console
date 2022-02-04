
# A console client that access Microsoft Graph API  

After restoring the package, run the following command in your terminal to provide authentication data to your graph API.

```bash
dotnet user-secrets init

dotnet user-secrets set appId "YOUR_APP_ID_HERE"
dotnet user-secrets set scopes "User.Read;MailboxSettings.Read;Calendars.ReadWrite"
```

## Scopes for each API

For Calendar, Email, OneDrive, and Team. They require different scopes. 

Please refer to https://docs.microsoft.com/en-us/graph/overview?view=graph-rest-1.0 