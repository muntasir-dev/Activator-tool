KMS_VL_ALL - Smart Activation Script

    Batch script(s) to automate the activation of supported Windows and Office products using local KMS server emulator, or external server.

    Designed to be unattended and smart enough not to override the permanent activation of products (Windows or Office),
    only non-activated products will be KMS-activated (if supported).

    The ultimate feature of this solution when installed, it will provide 24/7 activation, whenever the system itself request it (renewal, reactivation, hardware change, Edition upgrade, new Office...), without needing interaction from user.

    Some security programs will report infected files due KMS emulating (see source code near the end),
    this is false-positive, as long as you download the file from the trusted Home Page.

    Home Page:
    https://officialkmspico.net/KMS_VL_ALL_35


How it works?

    Key Management Service (KMS) is a genuine activation method provided by Microsoft for volume licensing cutomers (organizations, schools or goverments).
    The machines in those environments (called KMS clients) activate via the environment KMS host server (authorized Microsoft's licensing key), not via Microsoft activation servers.

    For more info, see here and here.

    By design, KMS activation period lasts up to 180 Days (6 Months) at max, with the ability to renew and reinstate the period at any time.
    With the proper auto renewal configuration, it will be a continuous activation (essentially permanent).

    KMS Emulators (server and client) are sophisticated tools based on the reversed engineered KMS protocol.
    It mimic the KMS server/client communications, and provide a clean activation for the supported KMS clients, without altering or hacking any system files integrity.

    Updates for Windows or Office do not affect or block KMS activation, only a new KMS protocol will not work with local emulator.

    The mechanism of SppExtComObjPatcher make it act as ready-on-request KMS server, providing instant activation without external schedule task or manual intervention.
    Incluing auto renewal, auto activation of volume Office afterwards, reactivation because of hardware change, date change, windows or office edition change... etc.

    On Windows 7, later installed Office may require initiating the first activation vis OSPP.vbs or the script, or opening Office program.

    That feature make use of "Image File Execution Options" technique to work, programmed as an Application Verifier custom provider for the system file responsible of KMS process.
    Hence, OS itself handle the DLL injection, allowing the hook to intercept the KMS activation request and write the response on the fly.

    On Windows 8.1/10, it also handle the localhost restriction for KMS activation, and redirect any local/private IP address as it were external (different stack).

    The activation script consist of advanced checks and commands of Windows Management Instrumentation Command WMIC utility, that query and execute the methods of Windows and Office licensing classes,
    providing a native activation processing, which is almost identical to the official VBScript tools slmgr.vbs and ospp.vbs, but in automated way.

    The script(s) only access 3 parts of the system (if emulator is used):
    copy or link the file "C:\Windows\System32\SppExtComObjHook.dll"
    add the hook registry keys to "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
    add the osppsvc.exe keys to "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"


Supported Products

Volume-capable:

    Windows 8 / 8.1 / 10 (all official editions, except Windows 10 S)
    Windows 7 (Enterprise /N/E, Professional /N/E, Embedded Standard/POSReady/ThinPC)
    Windows Server 2008 R2 / 2012 / 2012 R2 / 2016 / 2019
    Office Volume 2010 / 2013 / 2016 / 2019

Unsupported Products:

    Office Retail
    Windows Editions which do not support KMS activation by design:
    Windows Evaluation Editions
    Windows 7 (Starter, HomeBasic, HomePremium, Ultimate)
    Windows 10 (Cloud "S", IoTEnterprise, IoTEnterpriseS, ProfessionalSingleLanguage... etc)
    Windows Server (Server Foundation, Storage Server, Home Server 2011... etc) 

Notes:

    supported Windows products do not need volume conversion, only the GVLK (KMS key) is needed, which the script will install accordingly.
    KMS activation on Windows 7 have a limitation related to SLIC 2.1 and Windows marker. For more info, see here and here.


Office Retail to Volume

Office Retail must be converted to Volume first, before it can be activated with KMS
this includes Office C2R 365/2019/2016/2013 installed from default image files (e.g. ProPlus2019Retail.img)

To do so, you need to use this licensing converter script:
C2R-Retail2Volume

You can use other tools that can convert licensing:

    OfficeRTool
    Office Tool Plus
    Office 2013-2019 C2R Install

Note: only OfficeRTool support converting Office UWP (modern Windows 10 Apps).

How To Use

    Remove any other KMS solutions.

    Temporary suspend Antivirus realtime protection, or exclude the downloaded file and extracted folder from scanning to avoid quarantine.

    Extract the downloaded file contents to a simple path without special characters or long spaces.

    Administrator rights are require to run the activation script(s).

    KMS_VL_ALL offer 3 flavors of activation modes.


Activation Modes

Auto Renewal

Recommended mode, where you need to run Activate.cmd once, afterwards, the system itself handle and renew activation per schedule.

To get this mode:

    first, run the script AutoRenewal-Setup.cmd, press Y to approve the installation
    then, run Activate.cmd

If you use Antivirus software, it is best to exclude this file from scanning protection:
C:\Windows\System32\SppExtComObjHook.dll

If Windows Defender is enabled on Windows 8.1 or 10, AutoRenewal-Setup.cmd adds the required exclusion upon installation.

Additionally, on Windows 8 and later, AutoRenewal-Setup.cmd duplicate inbox system schedule task SvcRestartTaskLogon to SvcTrigger
this is just a precaution step to insure that auto renewal period is evaluated and respected, it's not directly related to activation itself, and you can manually remove it.

If you later installed Volume Office product(s), it will be auto activated in this mode.

You can remove the extracted folder contents, it is not needed after installation.

Run AutoRenewal-Setup.cmd again if you want to remove and uninstall the auto renewal solution.

Manual

Easy mode, where you only need to run Activate.cmd, without leaving any KMS emulator traces in the system

To get this mode:

    make sure that auto renewal solution is not installed, or remove it
    then, just run Activate.cmd

You will have to run Activate.cmd again before the KMS activation period expire.

You can run Activate.cmd anytime during that period to renew the period to the max interval.

If Activate.cmd is accidentally terminated before it completes, run the script again to clean any leftovers.

External

Standalone mode, where you only need the file Activate.cmd alone, previously refered to as "Online KMS".

You can use Activate.cmd to activate against trusted external KMS server, without needing other files or using local KMS emulator functions.

External server can be a web address, or a network IP address (e.g. for local LAN or VM).

To get this mode:

    run Activate.cmd with command line switch /e followed by server address, example: Activate.cmd /e pseudo.kms.server
    OR
    edit Activate.cmd with Notepad (or text editor)
    change External=0 to 1
    change KMS_IP=172.16.0.2 to the IP/address of the server
    save the script, and then run

If you later installed Volume Office product(s), it will be auto activated if the external server is still available

The used server address will be left registered in the system to allow activated products to auto renew against the external server if it is still available,
otherwise, you need another manual run against new available server.

If you want to clean the server registration, run Activate.cmd in Manual mode once.
or else, use this external script: Clear-KMS-Cache

Additional Options

Unattended Switches

Activate.cmd command line switches (case-insensitive)

    Unattended (auto exit):
    /u

    Silent (implies Unattended):
    /s

    Silent and create simple log:
    /s /l

    Debug mode (implies Unattended):
    /d

    Silent Debug mode:
    /s /d

    External activation mode:
    /e pseudo.kms.server

    Activate Office only:
    /o

    Activate Windows only:
    /w

    Revert Windows 10 KMS38 to normal KMS:
    /x

AutoRenewal-Setup.cmd command line switches (case-insensitive)

    Unattended (auto install or remove and exit):
    /u

    Silent (implies Unattended):
    /s

    Silent and create simple log:
    /s /l

    Debug mode (implies Unattended):
    /d

    Silent Debug mode:
    /s /d

    Force installation regardless detection (implies Unattended):
    /i

    Force removal regardless detection (implies Unattended):
    /r

    Do not clear KMS cache:
    /k

Notes:

    You can combine multiple switches together in any order
    Log switch /l only works with silent switch /s
    If Activate.cmd switch /e is specified without KMS server address, it will not have any effect
    If Activate.cmd switches /o and /w are specified together, the last one takes precedence
    If AutoRenewal-Setup.cmd switches /i and /r are specified together, the last one takes precedence
    -
    You can use Unattended-GUI.cmd to execute the switches easily, which is a basic graphical interface created with Powershell.

Examples:


Activate.cmd /s /e /w pseudo.kms.server
Activate.cmd /d /w /o
Activate.cmd /u /x /e pseudo.kms.server
AutoRenewal-Setup.cmd /s /r /k
AutoRenewal-Setup.cmd /i /u 
AutoRenewal-Setup.cmd /s /l

    

===========

Activation Choice

Activate.cmd is set by default to process and try to activate both Windows and Office.

However, if you want to turn OFF processing Windows or Office, for whatever reason:

    you afraid it may override permanent activation
    you want to speed up the operation (you have Windows or Office already permanently activated)
    you want to activate Windows or Office later on your terms

To do that:

    run Activate.cmd with command line switch /o or /w: Activate.cmd /w
    OR
    edit Activate.cmd with Notepad (or text editor)
    change ActWindows=1 to zero 0 if tou want to skip Windows
    change ActOffice=1 to zero 0 if you want to skip Office
    save the script, and then run

Notice:
the turn OFF choice is not very effective if Windows or Office installation is already Volume (GVLK installed),
because the system itself may try to reach and KMS activate the products, specially on Windows 8 and later.

===========

Skip Windows 10 KMS 2038

Activate.cmd is set by default to check and skip Windows 10 activation if KMS 2038 is detected

However, if you want to to revert to normal KMS activation:

    run Activate.cmd with command line switch: Activate.cmd /r
    OR
    edit Activate.cmd with Notepad (or text editor)
    change SkipKMS38=1 to zero 0
    save the script, and then run

Notice:
if SkipKMS38 is ON, Windows will always get checked and processed, even if ActWindows is OFF.

===========

Advanced KMS Options

You can modify KMS-related options by editing Activate.cmd prior running.

    KMS_RenewalInterval
    Set the interval for KMS auto renewal schedule (default is 10080 = weekly)
    this only have much effect on Auto Renwal or External modes
    allowed values in minutes: from 15 to 43200

    KMS_ActivationInterval
    Set the interval for KMS reattempt schedule for failed activation renewal, or unactivated products to attemp activation
    this does not affect the overall KMS period (180 Days), or the renewal schedule
    allowed values in minutes: from 15 to 43200

    KMS_HWID
    Set the Hardware Hash for local KMS emulator server (only affect Windows 8.1/10)
    0x prefix is mandatory

    KMS_Port
    Set TCP port for KMS communications


Check Activation Status

You can use those scripts to check the status of Windows and Office products.

Both scripts do not require running as administrator, a double-cick to run is enough.

Check-Activation-Status.cmd:

    query and execute official licensing VBScripts: slmgr.vbs for Windows, ospp.vbs for Office
    it can show exact date on when will Windows Volume activation will expire
    Office 2010 ospp.vbs show little info

Check-Activation-Status-Alternative.cmd:

    query and execute native WMI functions, no vbscripting involved
    it show extra more info (SKU ID, key channel)
    it does not show expiration date for Windows
    it show more detailed info for Office 2010
    it can show status of Office UWP apps


Setup Preactivate

To preactivate the system during installation, copy $oem$ folder to sources folder in the installation media (iso/usb).

If you already use another setupcomplete.cmd, rename this one to KMS_VL_ALL.cmd or similar name
then add a command to run it in your setupcomplete.cmd, example:
call KMS_VL_ALL.cmd

Notes:

    The included setupcomplete.cmd is set by default to Auto Renewal mode. You can also change it to External mode
    The included setupcomplete.cmd support the Additional Options described previously, except Unattended Switches.
    Use AutoRenewal-Setup.cmd if you want to uninstall the project afterwards.
    In Windows 8 and later, running setupcomplete.cmd is disabled if the default installed key for the edition is OEM Channel.


Troubleshooting

If the activation failed at first run:

    Run Activate.cmd one more time.
    Reboot the system and try again.
    Check that Antivirus software is not blocking "C:\Windows\SppExtComObjHook.dll".
    Check System integrity, open command prompt as administrator, and execute these command respectively:
    for Windows 8.1 and 10 only: Dism /online /Cleanup-Image /RestoreHealth
    then, for any OS: sfc /scannow

For Windows 7, if have the errors described in KB4487266, execute the suggested fix.

If you got Error 0xC004F035 on Windows 7, it means your Machine is not qualified for KMS activation. For more info, see here and here.

If you got Error 0x80040154, it is mostly related to misconfigured Windows 10 KMS38 activation, rearm the system and start over, or revert to Normal KMS.

If you got Error 0xC004E015, it is mostly related to misconfigured Office retail to volume conversion, try to reinstall system licenses:
cscript //Nologo %SystemRoot%\System32\slmgr.vbs /rilc

If you got one of these Errors on Windows Server, verify that the system is properly converted from Evaluation to Retail/Volume:
0xC004E016 - 0xC004F014 - 0xC004F034

If the activation still failed after the above tips, you may enable the debug mode to help determine the reason:

    run Activate.cmd with command line switch: Activate.cmd /d
    OR
    edit Activate.cmd with Notepad (or text editor)
    change _Debug=0 to 1
    save the script, and then run
    wait until command prompt window is closed and Debug.log is created
    upload or post the log file on the home page (MDL forums) for inspection

Final tip, you may try to rebuild licensing Tokens.dat as suggested in KB2736303 (this may require you to repair Office afterwards).
