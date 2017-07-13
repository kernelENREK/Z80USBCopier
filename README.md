# Z80USBCopier
Windows service for 'silent copy' plugged USB


## 1) What it does?
Z80USBCopier is a __Windows Service__ that automatically *'silent copy'* the content from a plugged USB to a HDD folder.

## 2) Give me a real scenario usage

**Scenario 1:**
Imagine someone said you: *hey dude, can you print me this ubercool.pdf i have on this USB?*

You plug your *'friend'* USB on your PC and instantly this shit start to make a 'silent copy' from USB content to your HDD.
Imagine your *'friend'* has this folder: *E:\Holidays\2016\Ibiza\Pacha\Hot_Pictures\wtf_0001.jpg,  E:\Holidays\2016\Ibiza\Pacha\Hot_Pictures\wtf_0002.jpg, E:\Holidays\2016\Ibiza\Pacha\Hot_Pictures\wtf_0003.jpg*, ... (hope you get the idea)

Without your *'friend'* knowing you copied all those (hot pictures?, really?) pictures to your HDD.

**Scenario 2:**
You admin a library, copy-shop, or similar business where people insert USB devices to make some actions.

Imagine 'person A' wants to print *E:\mycollegeHomeWork.doc* but he/she has this on the USB: *E:\someFolder\otherSubfolder\moreSubFolder\andMore\accountPasswords.txt* 

What do you think is *accountPasswords.txt*? :laughing:

**Scenario 3:**
You have **your own** USB and accidentally you delete or format the USB (or simply lost it, or the USB is damaged with a RAW format). You have a 'backup' on your HDD. :+1:

## 3) How it works?
Well, this is a Windows Service. No user interaction is needed.

![Alt text](https://i.imgur.com/1fcgkqu.png "Windows services")

The service behaviour is set-up by a [config.ini](https://github.com/kernelENREK/Z80USBCopier/blob/master/Z80USBCopier/Resources/config.ini) file

## 4) How long does it take to copy the whole USB to HDD?
Well, it depends on the USB size, if you copy all extensions or only a few ones or if you copy all file sizes or only some files of a certain size or if USB key / USB port is 3.0/3.1 compatible, ...

I have perfomed some test and with a very cheap USB key (think is an USB 1.0), copy 80 MB (196 files and 19 folders/subfolders) from USB to HDD (not SSD) takes around 20 seconds (Windows 10)

## 5) I've never used/debugged/coded a Windows Service. How this shit runs/install/or debug?
You can find useful information on [Z80USBSvc.vb](https://github.com/kernelENREK/Z80USBCopier/blob/master/Z80USBCopier/Z80USBSvc.vb) file

Basically you need to install service:
```
installutil Z80USBCopier.exe
```

and start the service:
```
net start Z80USBSvc
```
