# KeepOutlookRunning
Keep Outlook Running COM Addin Installable Without Admin Rights

### What Does This Do?
This Outlook add-in forces Outlook to minimize, instead of close, when you hit the "X" button. While minimized, it stays running and you won't miss out on important emails!

### How to Install:
1. Extract "KeepOutlookRunning.zip"
2. CD to folder containing the contents you just extracted
3. Run InstallKeepOutlookRunning.bat if you are running a 32 bit version of Outlook. 
Run InstallKeepOutlookRunning-64bit.bat if you are running a 64 bit version of Outlook
4. Restart Outlook

### How Does This Work?
This does not touch any system files. The .bat file installs the DLL to the current user's AppData folder, then manually installs all the necessary registry entries for the file to HKEY_CURRENT_USER and HKEY_CURRENT_USER\Software\Classes\ . By registering there, we avoid having to register in HKEY_CLASSES_ROOT which requires admin priveleges. I was able to find all the necessary registry entries for installation by generating a registy (.reg) file on the 32-bit dll with Visual Studio's regcap.exe, then in installation the batch file replace placeholders with the necessary local user paths.

Credit of course to Tim Eck who wrote the initial add-in for the StackOverflow Answer here: https://superuser.com/a/275244/322104

Enormous credit as well to Jens Huehn who's [answer](https://stackoverflow.com/a/55072381/3171509) was crucial to figuring it all out, and who's batch code I borrowed. His product [SlideFab 2](https://slidefab.com/lite/) served as the blueprint for making this work. 

The SourceCode is sourced from Ourselp's github here: https://github.com/Ourselp/KeepOutlookRunning
