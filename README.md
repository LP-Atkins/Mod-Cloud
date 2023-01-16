# Mod-Cloud Server Client, v1.05
A small app that syncs your server mods for various games via cloud services.
This application will only work for:
  1.) Windows
  2.) Steam games
  3.) URLs from onedrive (but you should be able to modify it to work with other platforms)

Written as a HTA (HTML Application) using Visual Basic Script.
Windows Only, launch using the HTA file, no external dependencies required.

## Version Summary:
* v1.00: Original release.
* v1.01: Converted to vbscript as windows defender was falsly flagging the .exe as a virus.
* v1.02: Added support for the steam overlay.
* v1.03: Changed the script to use a direct link to a zip file for mods, this allows the script to be a 'one click' operation for the end user.
* v1.04: Modified to support any game.
* v1.05: Packaged into a HTA for multi-game management with a GUI. (Renamed from Valheim-Cloud to Mod-Cloud)

  
## How to use:    ***NOTE*** The default setting in this file will need to have a modpack embedded URL added before using it.
    1.) Create a zip file of all your mods from the root directory of valheim.
    2.) Upload to onedrive and generate an embedded link.
    3.) Open the HTA file and add connection settings for each game.
    4.) Distribute the HTA file to clients.
    5.) Launch the game via the HTA file.
    
## What does this do exactly?:
    1.) Firstly it locates the game directory by looking at the registry using the steamID.
    2.) It will check if it has the latest modpack by getting the timestamp.
    3.) It will download the modpack if required.
    4.) Before it launches it will create a copy of the games current contents to an adjacent folder (should only have to do this once).
    5.) The modpack will be unpacked and then the game files patched.
    6.) The game is then launched with a steam command.
    7.) Once the application has been terminated the game will be unpatched back to the previous state, so remember to give your PC a few seconds before shutting it down as it is still unpatching the game files.

    
    Note1: Each time you launch a game you will need to launch it via this app.
    Note2: Check the 'How to use.pdf' file in the repository for more detailed instructions.
    Note3: If you want you can also modify the HTA file to pull a list of servers via the internet, that way you don't need to update the clients.... I might add this feature in future.
    Note4: If you accidentally shutdown the computer too quickly after exiting or had a powerfailure while running the game (or system crash) the game wont be unpatched, you can either browse to the game folder and restore the unpatched verion (steamapps\gamename(unmodded)) by deleting the game folder and renaming the unmodded copy.
