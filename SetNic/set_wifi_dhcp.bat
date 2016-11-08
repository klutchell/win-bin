:: Configuration Variables ::
set connectionName="Wi-Fi"

:: Change Nothing Below This Line ::
netsh interface ip set address %connectionName% dhcp
netsh interface ip set dns %connectionName% dhcp