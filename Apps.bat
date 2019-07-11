@echo off
color 0F
title Application Selector

echo.
echo *******************************************************************************
echo *                                                                             *
echo *                                                                             *
echo *                                                                             *
echo *                               Welcome Mr Cadd:                              *
echo *                                                                             *
echo *                                                                             *
echo *                                                                             *
echo *                                                                             *
echo *                                                                             *
echo *                                                                             *
echo *                               Today's Date is:                              *
echo *                                                                             *
echo *                                %date%                               *
echo *                                     %time:~0,5%                                   *
echo *                                                                             *
echo *******************************************************************************
echo.
pause

:ECHO
SET choice=''
SET choice1=''
cls
echo Application Selector
echo.
echo [1]                                                       Open All Applications
echo [2]                                                      Close All Applications
echo [3]                                                 Open A Specific Application
echo [4]                                         Close Snap-on Integration Assistant
echo [5]                                                             Open Teamviewer
echo [6]                                                            Close Teamviewer
echo [7]                                                                        Exit
echo.
set /p choice= Enter your choice: 
echo.

if '%choice%' =='' goto ECHO
if '%choice%' =='1' goto ALPHA
if '%choice%' =='2' goto BRAVO
if '%choice%' =='3' goto CHARLIE
if '%choice%' =='4' goto KILL
if '%choice%' =='5' goto TEAMVIEWER
if '%choice%' =='6' goto KILL2
if '%choice%' =='7' goto DELTA
goto ECHO

:ALPHA
start /d "C:\Program Files (x86)\Internet Explorer" iexplore.exe https://snaponepc.com/CStoneEPC/launch.sls
start /d "C:\Program Files (x86)\Internet Explorer" iexplore.exe https://dealerdaily.toyota.com/LogOn authn_try_count=0&contextType=external&username=string&contextValue=%2Foam&password=sercure_string&challenge_url=https%3A%2F%2FDealerdaily.toyota.com%2FLogOn&request_id=-321554163026219671&OAM_REQ=&locale=en_US&resource_url=https%253A%252F%252FDealerdaily.toyota.com%252F
start /d "C:\Program Files (x86)\Internet Explorer" iexplore.exe http://207.186.44.186/bin/start/wsStart.application
goto DELTA

:BRAVO
taskkill /f /im iexplore.exe /im wsStart_4.exe /im PQIntegrationAssistant.exe
goto DELTA

:KILL
taskkill /f /im PQIntegrationAssistant.exe
PING 1.1.1.1 -n 1 -w 2000>nul
goto DELTA

:KILL2
taskkill /f /im TeamViewer.exe
PING 1.1.1.1 -n 1 -w 2000>nul
goto DELTA

:TEAMVIEWER
pushd %~dp0
start "" cmd /c cscript Teamviewer.vbs
goto DELTA

:CHARLIE
cls
echo Application Selector
echo.
echo [1]                                                                 Snap-On EPC
echo [2]                                                                Dealer Daily
echo [3]                                                                         CDK
echo [4]                                                                 Service One
echo [5]                                                           Fisher Auto Parts
echo [6]                                                           Kunkel Auto Parts
echo [7]                                                                     Outlook
echo [8]                                                            Auto PartsBridge
echo [9]                                                                Parts Trader
echo [10]                                                                     Chrome
echo [11]                                                                  Paylocity
echo [12]                                                             BuyAnAccessory
echo [13]                                                                Accessories
echo [14]                                                                        TIS
echo [15]                                                                Trademotion
echo [16]                                                                 Calculator
echo [17]                                                              Certification
echo [18]                                                            MD Tire Express
echo [19]                                                                Parts Voice
echo [20]                                                                    Go Back
echo [21]                                                                       Exit
echo.
set /p choice1= Enter your choice: 
echo.

if '%choice1%' =='' goto CHARLIE
if '%choice1%' =='1' goto SNAPON
if '%choice1%' =='2' goto DEALERDAILY
if '%choice1%' =='3' goto CDK
if '%choice1%' =='4' goto SERVICEONE
if '%choice1%' =='5' goto FISHER
if '%choice1%' =='6' goto KUNKEL
if '%choice1%' =='7' goto OUTLOOK
if '%choice1%' =='8' goto AUTOPARTSBRIDGE
if '%choice1%' =='9' goto PARTSTRADER
if '%choice1%' =='10' goto CHROME
if '%choice1%' =='11' goto PAYLOCITY
if '%choice1%' =='12' goto BUYANACCESSORY
if '%choice1%' =='13' goto ACCESSORIES
if '%choice1%' =='14' goto TIS
if '%choice1%' =='15' goto TRADEMOTION
if '%choice1%' =='16' goto CALCULATOR
if '%choice1%' =='17' goto CERTIFICATION
if '%choice1%' =='18' goto MDTIRE
if '%choice1%' =='19' goto PARTSVOICE
if '%choice1%' =='20' goto ECHO
if '%choice1%' =='21' goto DELTA
goto CHARLIE

:SNAPON
pushd %~dp0
start "" cmd /c cscript Snapon.vbs
goto DELTA

:DEALERDAILY
pushd %~dp0
start "" cmd /c cscript DealerDaily.vbs
goto DELTA

:CDK
pushd %~dp0
start "" cmd /c cscript CDK.vbs
goto DELTA

:SERVICEONE
pushd %~dp0
start "" cmd /c cscript ServiceOne.vbs
goto DELTA

:FISHER
pushd %~dp0
start "" cmd /c cscript Fisher.vbs
goto DELTA

:KUNKEL
pushd %~dp0
start "" cmd /c cscript Kunkel.vbs
goto DELTA

:OUTLOOK
pushd %~dp0
start "" cmd /c cscript Outlook.vbs
goto DELTA

:AUTOPARTSBRIDGE
pushd %~dp0
start "" cmd /c cscript AutoPartsBridge.vbs
goto DELTA

:PARTSTRADER
pushd %~dp0
start "" cmd /c cscript PartsTrader.vbs
goto DELTA

:CHROME
start /d "C:\Users\jcadd\AppData\Local\Google\Chrome\Application" chrome.exe
goto DELTA

:PAYLOCITY
pushd %~dp0
start "" cmd /c cscript Paylocity.vbs
goto DELTA

:BUYANACCESSORY
pushd %~dp0
start "" cmd /c cscript BuyAnAccessory.vbs
goto DELTA

:ACCESSORIES
pushd %~dp0
start "" cmd /c cscript Accessories.vbs
goto DELTA

:TIS
pushd %~dp0
start "" cmd /c cscript Tis.vbs
goto DELTA

:TRADEMOTION
pushd %~dp0
start "" cmd /c cscript Trademotion.vbs
goto DELTA

:CALCULATOR
pushd %~dp0
start "" cmd /c cscript Calculator.vbs
goto DELTA

:CERTIFICATION
pushd %~dp0
start "" cmd /c cscript Certification.vbs
goto DELTA

:MDTIRE
pushd %~dp0
start "" cmd /c cscript MdTire.vbs
goto DELTA

:PARTSVOICE
pushd %~dp0
start "" cmd /c cscript PartsVoice.vbs
goto DELTA

:DELTA
cls
exit /b