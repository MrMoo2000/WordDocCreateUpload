# WordDocCreateUpload
This is a console app using .NET v7.0 that connects to your one drive using Microsoft Graph, using delegated access. From this app, you can:
- Browse your one drive directory and set an upload destination 
- Create a word document with a single run of text and upload it to your destination. 

## Configure
To configure this application you will need to register an application in Azure. You will need the following permissions with the type of Delegated:
- Files.ReadWrite
- User.Read

In the appsettings.json file, enter your client ID and tenant ID from the application overview in Azure. 
[Please see here for additional information on registing an app with Azure.](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app "Please see here for additional information on registing an app with Azure.")

##Purpose
The purpose of this project was to demonstrate using the Microsoft Graph API to create and upload a word document to OneDrive. 

[![Hack Together: Microsoft Graph and .NET](https://img.shields.io/badge/Microsoft%20-Hack--Together-orange?style=for-the-badge&logo=microsoft)](https://github.com/microsoft/hack-together) [![dotnet 7.0](https://img.shields.io/badge/Microsoft-.NET%207.0-blueviolet?style=for-the-badge&logo=dotnet)](https://dotnet.microsoft.com/) [![Microsoft Graph](https://img.shields.io/badge/Microsoft-%20Graph-orangered?style=for-the-badge&logo=Microsoft%20Office)](https://graph.microsoft.com)
[![Spectre.Console NuGet Version](https://img.shields.io/nuget/v/spectre.console.svg?style=flat&label=NuGet%3A%20Spectre.Console)](https://www.nuget.org/packages/spectre.console)
