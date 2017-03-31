# JuraScript
JuraScript is... a native .NET script interpreter with COM/ActiveX support, JavaScript console and an MSI installer to add support for the .JUR file extension.

JuraScript takes the [Jurassic](https://github.com/paulbartrum/jurassic) JavaScript engine and uses its API to extend functionality such that existing JScript scripts can be executed under Jurassic without any modifications. JuraScript is aimed at end-users, as well as IT administrators and web developers.

JuraScript's goals are to stay current with Jurassic and compatible with the latest version. Unfortunately, minor changes were needed to enhance JuraScript functionality. Requests have been made to have these changes added to the Jurassic engine.

# JuraScript has a nuget package now!
It lives here: [https://nuget.org/packages/JuraScript/1.0](https://nuget.org/packages/JuraScript/1.0)
You can use this command to install it from the gallery:<br/>
```PM> Install-Package JuraScript```

# Current Features:
* require() function and exports object from CommonJS
* COM/ActiveX Support - MS Office automation, ADO, WMI and more with the same syntax as Microsoft's JScript/WSH `var fso = new ActiveXObject("Scripting.FileSystemObject");`
* Windows MSI Installer
* Associates .JUR extension so that scripts can be executed by double-click or by typing the filename from the command-line (using PATHEXT environment variable at machine level)
* Partial legacy support for WSH objects such as `WScript.Echo()`, `WScript.StdIn.ReadLine()`, etc.

# JuraScript goals:
## Currently working on:
* Implement Active Scripting interfaces, allowing JuraScript to run under the Windows Scripting Host, Internet Explorer and Classic ASP environments
* Improved COM compatibility using conventions borrowed from WSH/JScript behavior
* Easy, step-by-step method of deployment for NT domains (both silent .MSI and script-based) for as many configurations as possible
* Offline Documentation
* Enhanced console functionality
## Potential future goals:
* Expose the full BCL to JuraScript
* Expose Assembly.Load() and allow JuraScript to instantiate late-bound .NET classes
* JuraScript ASP.NET MVC view
* Implement commonjs, commonjs unit tests
