# HP-UPD-Updater
Script to detect and upgrade HP Universal Print Driver (UPD) for PCL 6 and PostScript.  Designed for GPO deployment.  Can also perform silent fresh installs by running it with the -SilentInstall parameter

I got an email from HP warning me about critical security vulnerabilities in the UPD.
It linked to https://support.hp.com/us-en/document/ish_11892982-11893015-16/hpsbpi03995
Which is terrifying.  It lists several CVSS 9.8 vulnerabilities which have been there since the beginning of time.

I've poked at the UPD's install.exe command line parameters but can't find a combination that silently upgrades UPD.  I also found AutoUpgradeUPD.exe in hp's toolkit but it doesn't seem to actually do what the filename implies, so I created this powershell script to do HP's job.

Modern HP printers are garbage and would be more useful to society if they were burned for fuel.

Special thanks to Grok.com for helping me write this script.
