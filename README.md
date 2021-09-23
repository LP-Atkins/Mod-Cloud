# Valheim-Cloud Server Client, v1.00
A small app that syncs your server mods via cloud services.
This application is for Windows only.

Written in Visual Basic 6 (VB Classic)

## Version Summary:
* v1.00: Original release.

  
## How to use:    
    1. Ensure your valheim is in an unmodded 'vanilla' state.
    2. Using a cloud folder sharing application (I use onedrive) create a shared folder that all users can sync to their library. This folder will need to appear as a folder on the end users computer.
    3. Place this program into the folder and launch it.
    4. On first launch the program will create a mods directory and instruct you to load all of your mods into the folder before exiting.
    5. Place all of your server changes into the directory as if they were at the same level as the root directory of valheim (i.e BepInEx\Plugins\plugin.dll).
    6. Launch the program again and it will attempt to locate your valheim directory automatically and create a copy with the exe name appended. If it does not find your game files you will need to input the location manually, once it determines a valid path it will automatically do the rest.
    7. After the game files have been copied it will apply the server mods that are in the mods directory in the shared folder.
    8. Once complete the 'launch' button will be available to use.

    
    Note1: Each time you launch valheim you will need to launch it via this app.
    Note2: The 'Reload Files' button deletes the modded valheim directory, copies the vanilla one and patches the game again... this may be useful after a steam update.
