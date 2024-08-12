To use this macro, you need to open the message in edit mode. 
To test the macro, select a message containing large images and press Ctrl+C to copy the message then press Ctrl+V several times to make copies of the message, so you have plenty of messages to work with as you determine the best size for your needs.

1) Open the message.
2) Select the images(s) or entire message.
3) Run the macro.

** The picture size (in CM) is set in this line: picSize = 13 **

You'll need to set a reference to the Word Object Model in the VB Editor's Tools > References.

You need to have macro security set to the lowest setting, Enable all macros during testing. The macros will not work with the top two options that disable all macros or unsigned macros. You could choose the option Notification for all macros, then accept it each time you restart Outlook, however, because it's somewhat hard to sneak macros into Outlook (unlike in Word and Excel), allowing all macros is safe, especially during the testing phase. You can sign the macro when it is finished and change the macro security to notify.

To check your macro security in Outlook 2010 and newer, go to File, Options, Trust Center and open Trust Center Settings, and change the Macro Settings.

The macro should be placed in a module. 
Open the VBA Editor by pressing Alt+F11 on your keyboard.

To put the code in a module:
- Right click on Project1 and choose Insert > Module
- Copy and paste the macro into the new module.

Set a reference to other Object Libraries
If you receive a "User-defined type not defined" error, you need to set a reference to another object library.
1) Go to Tools, References menu.
2) Locate the object library in the list and add a check mark to it. (Word and Excel object libraries version numbers will match Outlook's version number.)
