Cloud VBA
---------
Two-way communication between Word VBA and HTML5

Prerequisites:
1) Install Word 2013 or later
2) Enable macros in Word

Demo:
1) Identify the local location of your messenger.docm. It's called localLocation in point 2.
2) Access the following URL from your browser: 'ms-word:ofe|u|' + localLocation
    eg. ms-word:ofe|u|C:\GitHub\cloud-vba\cloud.docm
3) Allow Word to open the cloud.docm document.
4) Answer the prompt.
5) Notice a web page open with your input.
6) This completes the 2-way communication demo, but in practice the message to web will usually trigger a localStorage even for cooler use cases.
