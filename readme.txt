Cloud VBA
---------
Two-way communication between Word VBA and HTML5

Prerequisites:
1) Install Word 2013 or later
2) Enable macros in Word
3) NodeJS (or any server, the UI is agnostic, but I used Express)

Demo:
1) In the command line, run "npm install" then "node app" to start the lightweight file server.
2) Identify the local location of your messenger.docm. It's called localLocation in point 2.
3) Access the following URL from your browser: ms-word:ofe|u|http://localhost:3000/files/cloud.docm?action=PromptUserForInput
    Note: The reason we need a file server is because the protocol requires a remote URL. Local locations won't invoke the 
          The protocol refers to ms-word:ofe|u|http://, as in window.location.protocol
          The ms-word part by itself is called the scheme name
          This approach leverages .\Microsoft Office\...\protocolhandler.exe via a registry entry to open a document by URL
    Learn More: https://msdn.microsoft.com/en-us/library/office/dn906146.aspx
3) Allow Word to open the cloud.docm document.
4) Answer the prompt.
5) Notice a web page open with your input.

That completes the 2-way communication demo, but in practice you can do much more than was shown.
One nifty pattern is to have the message trigger a localStorage event by writing listeners in your larger JS app.
Such communication can be done in the background without user interaction as well.
You can hide the Word splash if you use a custom protocol hooked to a custom registry entry. That entry can call Word in silent mode.

Tech notes and todos:
  - ref: https://expressjs.com/en/starter/hello-world.html
