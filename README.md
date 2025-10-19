# KeepOutlookRunning
Keep Outlook Running COM Addin (Made Installable Without Admin Privileges)

### What Does This Do?
This Outlook add-in forces Outlook to minimize, instead of close, when you hit the "X" button. While minimized, it stays running and you won't miss out on important emails! The batch file here allows users without Admin Priveleges to install this add-in, such as those working on a company computer.

### How to Install:
1. Extract contents of "KeepOutlookRunning.zip" anywhere
2. CD to folder containing the contents you just extracted using CMD or Power Shell
3. Run InstallKeepOutlookRunning.bat if you are running a 32 bit version of Outlook, 
Run InstallKeepOutlookRunning-64bit.bat if you are running a 64 bit version of Outlook
4. Restart Outlook

**Note:** You may have to change some of the paths in the .reg file depending on your version of Outlook. Right now it is configured for Outlook 2016, the only one I have to test it on.

### How Does This Work?
This does not touch any system files. The .bat file installs the DLL to the current user's AppData folder, then automatically inputs all the necessary registry entries for the file to HKEY_CURRENT_USER and HKEY_CURRENT_USER\Software\Classes\ . By registering there, we avoid having to register in HKEY_CLASSES_ROOT which requires admin privileges. 

How was I able to find all the necessary registry entries? I was able to find all the necessary registry entries for installation by generating a registy (.reg) file on the 32-bit dll with Visual Studio's regcap.exe. The installation batch file replaces installation path placeholders with the relevant local user paths.

### Tested and Confirmed Working On
- Microsoft® Outlook® for Microsoft 365 MSO (Version 2210 Build 16.0.15726.20070) 64-bit
- Apparently working on newer versions as well

### Credit
Credit of course to Tim Eck who wrote the initial add-in for the StackOverflow Answer here: https://superuser.com/a/275244/322104

Enormous credit as well to Jens Huehn who's [answer](https://stackoverflow.com/a/55072381/3171509) was crucial to figuring it all out, and who's batch code I borrowed. His product [SlideFab 2](https://slidefab.com/lite/) served as the blueprint for making this work without Admin Priveleges.

The SourceCode is sourced from Ourselp's github here: https://github.com/Ourselp/KeepOutlookRunning
