@echo off
if "%IsEmulated%"=="true" goto :eof
start /wait msiexec /quiet /i %RoleRoot%\plugins\RemoteForwarder\RemoteForwarder.msi

@echo Checking firewall rule

netsh advfirewall firewall show rule name="WaRemoteForwarderService rule" 

if ERRORLEVEL 1 (
	@echo Adding firewall rule for remote forwarder

        netsh advfirewall firewall add rule name="WaRemoteForwarderService rule" description="Allow incoming connections to the forwarder" dir=in protocol=tcp program="%ProgramFiles%\Windows Azure Remote Forwarder\RemoteForwarder\RemoteForwarderService.exe" action=allow enable=yes
)

