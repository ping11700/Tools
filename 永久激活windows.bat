%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

slmgr /skms one.kmspass.com && slmgr /ato

@REM 适用版本：
@REM   1 = Windows Server 2008 Standard                    63 = Windows 7 Professional N
@REM   2 = Windows Server 2008 Standard V                  64 = Windows Vista Enterprise
@REM   3 = Windows Server 2008 Datacenter                  65 = Windows Vista Enterprise N
@REM   4 = Windows Server 2008 Datacenter V                66 = Windows Vista Business
@REM   5 = Windows Server 2008 Enterprise                  67 = Windows Vista Business N
@REM   6 = Windows Server 2008 Enterprise V                68 = Windows ThinPC
@REM   7 = Windows Server 2008 Web                         69 = Windows Embedded POSReady 7
@REM   8 = Windows Server 2008 Compute Cluster             70 = Windows Embedded Industry 8.1
@REM   9 = Windows Server 2008 R2 Standard                 71 = Windows Embedded Industry E 8.1
@REM  10 = Windows Server 2008 R2 Datacenter               72 = Windows Embedded Industry A 8.1
@REM  11 = Windows Server 2008 R2 Enterprise               73 = Office Access 2010
@REM  12 = Windows Server 2008 R2 Web                      74 = Office Excel 2010
@REM  13 = Windows Server 2008 R2 Compute Cluster          75 = Office Groove 2010
@REM  14 = Windows Server 2012 Datacenter                  76 = Office InfoPath 2010
@REM  15 = Windows Server 2012 Standard                    77 = Office Mondo 2010
@REM  16 = Windows Server 2012 MultiPoint Premium          78 = Office Mondo 2010
@REM  17 = Windows Server 2012 MultiPoint Standard         79 = Office OneNote 2010
@REM  18 = Windows Server 2012 R2 Datacenter               80 = Office OutLook 2010
@REM  19 = Windows Server 2012 R2 Standard                 81 = Office PowerPoint 2010
@REM  20 = Windows Server 2012 R2 Cloud Storage            82 = Office Project Pro 2010
@REM  21 = Windows Server 2012 R2 Essentials               83 = Office Project Standard 2010
@REM  22 = Windows Server 2016 Datacenter Preview          84 = Office Publisher 2010
@REM  23 = Windows 10 Enterprise                           85 = Office Visio Premium 2010
@REM  24 = Windows 10 Enterprise N                         86 = Office Visio Pro 2010
@REM  25 = Windows 10 Enterprise LTSB                      87 = Office Visio Standard 2010
@REM  26 = Windows 10 Enterprise LTSB N                    88 = Office Word 2010
@REM  27 = Windows 10 Education                            89 = Office Professional Plus 2010
@REM  28 = Windows 10 Education N                          90 = Office Standard 2010
@REM  29 = Windows 10 Professional                         91 = Office Small Business Basics 2010
@REM  30 = Windows 10 Professional N                       92 = Office Access 2013
@REM  31 = Windows 10 Home                                 93 = Office Excel 2013
@REM  32 = Windows 10 Home N                               94 = Office InfoPath 2013
@REM  33 = Windows 10 Home Single Language                 95 = Office Lync 2013
@REM  34 = Windows 10 Home Country Specific                96 = Office Mondo 2013
@REM  35 = Windows 8.1 Enterprise                          97 = Office OneNote 2013
@REM  36 = Windows 8.1 Enterprise N                        98 = Office OutLook 2013
@REM  37 = Windows 8.1 Professional WMC                    99 = Office PowerPoint 2013
@REM  38 = Windows 8.1 Professional                       100 = Office Project Pro 2013
@REM  39 = Windows 8.1 Professional N                     101 = Office Project Standard 2013
@REM  40 = Windows 8.1 Core                               102 = Office Publisher 2013
@REM  41 = Windows 8.1 Core N                             103 = Office Visio Standard 2013
@REM  42 = Windows 8.1 Core ARM                           104 = Office Visio Pro 2013
@REM  43 = Windows 8.1 Core Single Language               105 = Office Word 2013
@REM  44 = Windows 8.1 Core Country Specific              106 = Office Professional Plus 2013
@REM  45 = Windows 8.1 Core Connected                     107 = Office Standard 2013
@REM  46 = Windows 8.1 Core Connected N                   108 = Office Professional Plus 2016
@REM  47 = Windows 8.1 Core Connected Single Language     109 = Office Project Pro 2016
@REM  48 = Windows 8.1 Core Connected Country Specific    110 = Office Visio Pro 2016
@REM  49 = Windows 8.1 Professional Student               111 = Office Publisher 2016
@REM  50 = Windows 8.1 Professional Student N             112 = Office Access 2016
@REM  51 = Windows 8 Professional WMC                     113 = Office Skype for Business 2016
@REM  52 = Windows 8 Professional                         114 = Office Mondo 2016
@REM  53 = Windows 8 Professional N                       115 = Office Visio Standard 2016
@REM  54 = Windows 8 Enterprise                           116 = Office Word 2016
@REM  55 = Windows 8 Enterprise N                         117 = Office Excel 2016
@REM  56 = Windows 8 Core                                 118 = Office Powerpoint 2016
@REM  57 = Windows 8 Core N                               119 = Office OneNote 2016
@REM  58 = Windows 8 Core Country Specific                120 = Office Project Standard 2016
@REM  59 = Windows 8 Core Single Language                 121 = Office Standard 2016
@REM  60 = Windows 7 Enterprise                           122 = Office Mondo R 2016
@REM  61 = Windows 7 Enterprise N                         123 = Office Outlook 2016
@REM  62 = Windows 7 Professional
