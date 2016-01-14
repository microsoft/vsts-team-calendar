# Team Calendar Extension 

Team Calendar is an extension for Visual Studio Online that demonstrates one type of experience **you** can now develop and install into your Visual Studio Online account. Eventually you will be able to share your extensions with other users.
To learn more about Extensions, see the [overview of extensions](https://www.visualstudio.com/en-us/integrate/extensions/overview).

**Note:** Extensions is currently in private preview. To get early access, join the free [Visual Studio Partner Program](http://www.vsipprogram.com/join).

## Overview

Extensions enable you to create first-class integration experiences within Visual Studio Online, just the way you have always wanted. An extension can be a simple context menu or toolbar action or can be a complex and powerful custom UI experience that light up within the account, collection, or project hubs. 

The Team Calendar extension provides Visual Studio Online teams an integrated calendar view in web access:

![screenshot](images/calendar-screen-shot.png)

### What the calendar shows

See [overview](overview.md) to learn more about the features of this extension.

## How to add new event sources

The Team Calendar extension is designed to be extended by other extensions. Other extensions can contribute new "event sources", which will be pulled from when the Team Calendar is rendered. Once you develop your extension, install it in the account that you installed the Team Calendar extension into.

See the [public-events sample](https://github.com/Microsoft/vso-extension-samples/tree/master/public-events) for an example of an extension that contributes to the Team Calendar.
