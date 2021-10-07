**This is deprecated and only works for Microsoft Outlook for Mac versions <16.16.20**

# OutlookURLHandler

Enables `outlook://msg-id` style links anywhere on OS X.

### How to use the custom URL handler?

1. Download the [`OutlookURLHandler.app`](https://github.com/acidix/OutlookURLHandler/releases/download/0.1/OutlookURLHandler.app.zip) file and place it into `/Applications`
2. Run the application once for it to register the URL Handler
3. You're all set for `outlook://msg-id` style links!

To be able to quickly and easily generate the links and put them in the clipboard, I created a second script, `MsgLinkToClpbrd` which puts the message link that calls the handler in your clipboard.

1. Download [`MsgLinkToClpbrd.scpt`](https://github.com/acidix/OutlookURLHandler/releases/download/0.1/OutlookURLHandler.scpt) 
2. Create the folder structure `/Users/your-user/Library/Scripts/Applications/Microsoft/Outlook` and place the script there
3. Use BetterTouchTool, FastScripts or any other tool to assign a hotkey to the script execution from within Outlook.

Now you can use direct links to your emails from anywhere on your Mac, with simple, plain-text links in the `outlook://msg-id` notation.
