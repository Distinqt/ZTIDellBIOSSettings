ZTIDellBIOSSettings
===================

The go to script for zero-touch Dell BIOS configuration.

# What does ZTIDellBIOSSettings do?
ZTIDellBIOSSettings (what a mouthful) is a vbs script that can be included in a MDT or SCCM+MDT environment. It does several things.  

It sets up the CCTK (Client Configuration ToolKit) environment, initializes WMI, installs the HAPI driver, and checks all the files for existence.

ZTIDellBIOSSettings allows you to do the following:
- Pass arguments to CCTK  
```
cscript ZTIDellBIOSSettings.wsf /arguments:"--assettag=PC1000"
```
- Give CCTK a settings file  
```
cscript ZTIDellBIOSSettings.wsf /settings:"%path to file%"
```
- Order CCTK to change the password  
```
cscript ZTIDellBIOSSettings.wsf /password
```
- Or of course a combination of the above

You are probably thinking by now, wait a minute. So all the work gets forwarded to CCTK?
That's a useless 250+ lines script.

You are almost right, I did not posses the will (or time for that matter) to reverse enigneer CCTK and incorporate this into my script, what I did do was create an awesome convenience wrapper for CCTK.

For instance this script has a great password tester. When you give ZTIDellBIOSSettings an argument it will try to find the right password for the BIOS in the _BIOSPasswords_ list variable. Once it has found a match it adds that to the commandline it gives to CCTK.  
This has the huge added advantage of not having to hope all your systems have the same BIOS password. CCTK itself has no way of doing this or even testing if there is a password required.

Via the _BIOSArguments_ list variable or the "Arguments" argument you could do all kind of dynamic things.
Like set the assettag or propertyownertag of the bios to the computername.

It has helpful logging and errorhandling to pinpoint problems with the script or CCTK.
It even translates CCTK errorcodes with the help of "cctkerrorcodes.xml" into a human readable error message!

__What are you waiting for? Download it and give it a spin.__

# Used locations and variables
ZTIDellBIOSSettings uses all kinds of variables or locations to get the information it needs.  
Below you will find a list.


__CCTK files__  
ZTIDellBIOSSettings requires CCTK files to function correctly, see the "Required files" section for more info on how to get and prepare those files.  

There are two places the script looks for the files:  

1. In the path found in the _CCTKPath_ variable.  
2. It searches the root of the Boot Image for the folder "CCTK".
3. A relative path can also be provided in the _CCTKPath_ variable. It will then use that relative path + the current working directory to try and hopfully locate the required CCTK files. (This is enormously useful in SCCM packages.)

The root (X:) could look like this:

>Program Files  
Users  
Windows  
..  
CCTK  

Or:
>Program Files  
Users  
Windows  
_SMSTaskSequence\Packages\xxxxxxxx\CCTK  
with the _CCTKPath_ variable like:
```
CCTK\
``` 

__Arguments__  
Soon

__Settings files__  
Soon

__Password list__  
Soon

__New password__  
Soon

#Required files
ZTIDellBIOSSettings requires the "ZTIUtility.vbs" file found in the MDT deployment share "script" folder (MDT) or at the same place in the MDT package (SCCM).  
If you don't want to use MDT you could get by with only using "ZTIUtility.vbs".  

Furthermore ZTIDellBIOSSettings needs some CCTK files. So grab a copy of CCTK here:  
http://en.community.dell.com/techcenter/systems-management/w/wiki/1952.dell-client-configuration-toolkit-cctk.aspx  

Once you have downloaded CCTK, follow these steps:  

1. Install the toolkit.
2. Go to the install folder of CCTK. (default is: "C:\Program Files (x86)\Dell\CCTK").
3. If you want human readable errors, copy "cctkerrorcodes.xml", otherwise skip to step 5.
4. Paste it into the folders: "X86" and "X86_64".
5. Create 2 new folders named "Prestart-x86" and "Prestart-x64" in your sources folder (SCCM) or on your deployment share (MDT).
6. In each of the previously created folders create a folder named "CCTK".
7. Now from the install folder of CCTK copy the insides of the folder "X86" into "Prestart-x86\CCTK\" and do the same for "X86_64", this ofcourse into "Prestart-x64\CCTK".
8. Finally you have to get this folder into the SCCM or MDT "Boot Image". The easiest way to do this is to reference this folder in the "Extra Files"(MDT) or "Prestart"(SCCM) boot image option. Here you can select the folder which is name contains the same architecture as your boot image is (Prestart-x64 for a x64 boot image).

If you add the CCTK files some other way you can supply this custom path via the _CCTKPath_ variable.
The value could be something like:  
```
\\server\share\CCTK\%Architecture%\
```
or
```
X:\ImageXAddedFolder\CCTK
```

#Setup
soon

# FAQ
__What is CCTK?__  
>It stands for __Client Configuration Toolkit__, it is developed by Dell for their desktops and laptops.  
It allows you to change BIOS/UEFI settings for Dell machines from within Windows(PE).  

__Why not just use CCTK to generate an .exe of the settings for me?__
>__First__ of, you would miss out on the dynamic nature that a script brings. You could give ZTIDellBIOSSettings a variable to use in the settings, this for instance allows you to set the assettag on a per machine basis.  
Like so:  
``cscript ZTIDellBIOSSettings.wsf /arguments:"--assettag=%OSDCOMPUTERNAME%"``  

>__Second__, you could easily add a new settings file or change an existing one. Because you can place all of them on a share or in a SCCM package.

>__Third__, you will not have the awesome password tester this script has. When you give ZTIDellBIOSSettings an argument it will try to find the right password for the BIOS in the _BIOSPasswords_ list variable. And once found, add that password to the commandline it gives to CCTK.
This way you also have your passwords seperated from your commandlines, which is a whole lot safer.

__Will this script work on other brands?__
>__No__, currently it is specifically made for Dell.   
It will even halt and throw an error if it detects something other than "Dell Inc."  
So if you want to use your task sequence on all your machines. You should either, guard the script with a condition on %MAKE% == "Dell Inc." or let the script continue on error.
I would suggest the former, if you let it continue on error you will not know if it succeded or not. It could have failed on an unknown password or non existing settings file. 

__When should I use the BIOSArgument list variable instead of the "Arguments" argument?__
>If you have more than one argument that contains spaces.
e.g.
```
cscript ZTIDellBIOSSettings.wsf /arguments:"--assettag="this requires quotes due to spaces" --propowntag="this too""
```
Because of the way the commandline argument parser of VBScript works it will strip all quotes but the outer ones. It leaves the command line like this:
```
cscript ZTIDellBIOSSettings.wsf /arguments:"--assettag=this requires quotes due to spaces --propowntag=this too"
```
Afterwards you have the assettag filled with:
```
this requires quotes due to spaces --propowntag=this too
```
And an empty propowntag.

>If you use the list, the script uses "ZTIUtility.vbs" which doesn't suffer from this. 

-----------------
The code is provided as-is and I hold no responsability.

You are allowed to use this for non-commercial usage.  
Contact me for commercial usage, commercial usage is, but not limited to, distribution of these source files, integration into commercial software, or using it as content for your commercial services.

If you would like to use excerpts for an article or book, please feel free to do so but do contact me, I would love to have a chat.

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/eeb9f83085564f822d849a2c34a0903f "githalytics.com")](http://githalytics.com/NinoFloris/ZTIDellBIOSSettings)
