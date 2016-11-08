:: Configuration Variables ::
set connectionName="Ethernet"

:: Change Nothing Below This Line ::
netsh interface ip set address %connectionName% dhcp
netsh interface ip set dns %connectionName% dhcp