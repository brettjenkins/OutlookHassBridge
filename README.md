# OutlookHassBridge

This is a very simple C# program that adds as an VSTO addon into outlook and updates an endpoint with a simple bool showing if there are unread emails or not as emails come in and are read.

It will PUT a JSON formatted version of OutlookStatus to an endpoint of your choice. You'll need to edit the URL in App.config.

I use this to PUT my unread email status to NodeRed, which then gets parsed into HASS - which explains the name of this project. But there is no reason you can't use for other uses.