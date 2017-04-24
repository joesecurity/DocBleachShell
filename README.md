# DocBleachShell

DocBleachShell is the integration of the great [DocBleach](https://github.com/docbleach/DocBleach) Content Disarm and Reconstruction tool into the Microsoft Windows Shell Handler.

**By using DocBleachShell documents are automatically disarmed before they are opened by Microsoft Word, Excel or Powerpoint.** As a result end users who have installed DocBleachShell are protected from exploits and malicious macros. DocBleachShell also comes with a [Joe Sandbox Cloud](https://www.joesecurity.org/joe-sandbox-cloud) integration. **Successfully bleached documents are automatically analyzed by Joe Sandbox Cloud**. In Joe Sandbox Cloud users can enable alerts, e.g. automated emails on detection of malicious files. With that CERTs, CIRTS or SOCs are automatically notified if their users where attacked by malicious documents. 

![DocBleacShell Overview](https://raw.githubusercontent.com/joesecurity/docbleachshell/master/img/shell.png)

DocBleachShell modifies the Microsoft Windows Shell Handler via the HKEY_CLASSES_ROOT\Type\shell\open\command registry key. 

# License

Code is developed in C# / .Net 4 and licensed under MIT. 

# Requirements

* Latest [Microsoft .Net Framework](https://www.microsoft.com/en-us/download/details.aspx?id=53344)
* Latest [Java Runtime Environment](https://java.com/de/download/), make sure that your Java bin directory is part of the PATH environment variable.

# Installation

To install DocBleachShell, call **DocBleachShell.exe -install** from an Administrator shell:

![DocBleacShell Overview](https://raw.githubusercontent.com/joesecurity/docbleachshell/master/img/install.png)

DocBleachShell will search and replace all shell handlers for Word, Excel and Powerpoint. For each registry modification a backup is made to the backup folder. After installation do not move DocBleachShell to any other path. To uninstall DocBleachShell call **DocBleachShell.exe -uninstall**:

![DocBleacShell Overview](https://raw.githubusercontent.com/joesecurity/docbleachshell/master/img/uninstall.png)

Once installed, if you open an Office file, DocBleachShell will be started. DocBleachShell will then call DocBleach which will disarm the document. Finally DocBleachShell will start Office to open the disarmed file.

# Logging

DocBleachShell uses log4net. The log file is located in the main directory and named "DocBleachShell.log".

# Configuration

Configuration of DocBleachShell is controlled via DocBleachShell.exe.config:

```xml
  <appSettings>
     <add key="OnlyBleachInternetFiles" value="true"/>
     ...
  </appSettings>
```

By default only documents downloaded from the Internet are bleached. This is done via the NTSF ADS Zone.Identifier check. You can turn off this check and bleach any document. To do so change the config "OnlyBleachInternetFiles" to false.

# Joe Sandbox Cloud Integration

DocBleachShell offers integration of [Joe Sandbox Cloud](https://www.joesecurity.org/joe-sandbox-cloud). Joe Sandbox Cloud enables to deeply analyze and detect malicious files. The Joe Sandbox Cloud integration can be enabled via DocBleachShell.exe.config, by adding your API Key:

```xml
  <appSettings>
     ...
     <add key="JoeSandboxCloudAPIKey" value="addyourapikeyhere"/>
  </appSettings>
```
After that, DocBleachShell will upload any document which DocBleach has disarmed. If the document is safe (i.e. DocBleach did not do any disarming) the document is not uploaded. 

# Links

* [DocBleach](https://github.com/docbleach/DocBleach) 
* [Joe Sandbox Cloud](https://www.joesecurity.org/joe-sandbox-cloud).

# Author

Joe Security (@[joe4security](https://twitter.com/#!/joe4security) - [webpage](https://www.joesecurity.org))

# Credits

Kudos to [PunKeel](https://github.com/PunKeel) for the very cool DocBleach project!
