 

# context

## [Office](Office.html). context

The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [documentation on MSDN](https://msdn.microsoft.com/EN-US/library/office/fp161104.aspx).

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.html)| 1.0|
|Applicable Outlook mode| Compose or read|

### Namespaces

<dl>

<dt>[mailbox](Office.context.mailbox.html)</dt>

<dd class="ms-font-m">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</dd>

</dl>

### Members

####  displayLanguage :String

Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.

The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.html)| 1.0|
|Applicable Outlook mode| Compose or read|

##### Example

```
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message; 
}
```

####  roamingSettings :[RoamingSettings](RoamingSettings.html)

Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.

The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.

##### Type:

*   [RoamingSettings](RoamingSettings.html)

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.html)| 1.0|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|