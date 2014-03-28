# CustomOffice365OWA
## A Custom Outlook Web App for Office 365

This is a Proof-of-Concept type of app for testing out/playing with the REST APIs for Office 365.

For some background info on this check out the following blog posts:  
http://mobilitydojo.net/2014/03/24/microsoft-provides-a-restful-api-for-exchange-part-1/  
http://mobilitydojo.net/2014/03/25/microsoft-provides-a-restful-api-for-exchange-part-2/  

### Getting Started
To test the code you will need to have an Office 365 account to retrieve data from. It will not work with on-premises versions of Exchange.

You should also have an Azure subscription to be able to configure settings for your app. (The app itself does not have to run on Azure, but there needs to be an entry in Azure AD.) 
The Azure AD features you will use does not require any paid services, so there should be no direct costs attached to setting it up from the developer side.

The Azure tenant you use to develop the app, and the Office 365 subscription you test with does not have to be related in any way. (Which is part of the idea for the new APIs.)

There are two lines of code you need to change to get the code to run (assuming you have the requirements above in place):  
_Web.config_

```
<add key="ida:ClientID" value="Retrieve this value from the Azure Management Portal." />
<add key="ida:Password" value="Retrieve this value from the Azure Management Portal." />
```

These settings can be found on the "Configure" tab of your application in the Azure Management Portal.

_ClientID_ is the "Client ID" guid on the page.  
_Password_ is one of the "keys" specified. (You can create multiple keys.)