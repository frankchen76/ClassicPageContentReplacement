# Getting Started with Create React App

This project demonstrate how to replace a content for a CEWP web part for SPO classic pages

## Instructure

* Create a credential store from your control panel. 
* Use the credential name in the code.
```C#
        private static ClientContext GetClientContext(string url, string credentialStoreName = "[CredStoreName]")
        {
            SecureString se = new SecureString();
            Credential cred = new Credential() { Target = "[CredStoreName]" };
            ClientContext ret = new ClientContext(url);
            cred.Load();
            ret.Credentials = new SharePointOnlineCredentials(cred.Username, cred.SecurePassword);
            return ret;
        }

```

