Create on a removable drive AUTORUN.INF file that you can't remove or rename the standard means, except formatting, thereby preventing it from infecting.
Support FAT/FAT32 only.

Command line parameters:
[/command [/mode:x]]

command:
/drives:ABCZ - vaccinate drives (A-Z);
/drives+ - vaccinate all drives;
/system+ - computer vaccination;
/system- - remove computer vaccination;
/resident - start program hidden and prompt for vaccinating every new drive;
/q - quit from resident app.

/mode:0 - hide all messages;
/mode:1 - show warning only messages (default);
/mode:2 - show all messages.
