OfficeJS-Vandivier
==============
Objectives
--------------
- Demonstrate usage of the Javascript API for Office, aka Office.JS.
- In the context of an MS Word Task Pane application.

Dependencies
------------
  - MS Word 2013 or later
  - NodeJS

Installation
------------
- From the command line, "npm start"
- Create a network shared folder from .\office-js-demo\public
  - Right click the folder and go to properties -> sharing, click share, and confirm sharing
- Copy the shared location, paste it in a text file, and transform it as follows:
  - "file://DESKTOP-MMSRKCH/public" should become "\\DESKTOP-MMSRKCH\public"
- Open .\public\files\office-js-demo.docm
- Navigate to file -> options -> trust center -> trust center settings
- Ensure macros are fully enabled and do not open documents in protected mode
- Navigate further to trust center settings -> trusted app catalog
- Add the shared folder path as a trusted app catalog and check "show in menu" on the right
- Close the docm and exit Word completely, to ensure the settings take effect
- When you open the docm again you should see a message that the XML listener has instantiated
- Go to the insert tab, click the 'My Add-ins' button, 'shared folder' in the dialog, refresh, and finally double click "OfficeJS Demo"
- The task pane app is inserted on the right. Have fun!


Demos
-----
- Doc Inject: Injection type can be text, html, or ooxml. It will be injected into the document.
- Like doc inject, but injects into the bottom of the task pane.
- Create Text Binding: Creates a content control w/ specified ID.
- Nav track by id: Listens for clicks into a content control specified by ID.
  - Watch the results area at task pane bottom for a click counter
- Change track by id: Listens for content changes in a control specified by ID.
- Update Bound Field through Log Bindings do what they sound like.
- Goto by ID: Put a content control id in the field above and click. It will take your cursor there.
- Send Data to VB: By far the coolest thing that will be demo'd.
  - Type "hello" and click the button. You just triggered a VBA macro from the task pane!