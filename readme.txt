Cloud VBA
---------
Two-way communication between Word VBA and HTML5

Prerequisites:
1) Word 2013 or later.
2) NodeJS (UI is agnostic, but Express Node is configured for the packaged demo)
3) Chrome (Easy change to work w/ others, not currently set up)

Demo:
1) In the command line, "npm install" then "node app"
2) Add a popup blocker exception for http://localhost:3000/ in Chrome, then open that URL
3) Click the button to Authorize Handler. If prompted, confirm Word is allowed to open documents from web.
3) Click the button to Request Word Input. Allow Word to open the document with macros.
4) Answer the prompt and confirm.
5) Notice the input was added to the web page.

That completes the 2-way communication demo.
You can hide the Word splash if you use a custom protocol hooked to a custom registry entry. That entry can call Word in silent mode.

Troubleshooting:
  - If you have issues activating macros, go to Word -> File -> Options -> Trust Center -> Trust Center Settings
  - Lower the macro security settings and unset the preference to open documents in protected mode.

Tech notes and todos:
  - ref: https://expressjs.com/en/starter/hello-world.html
