:: Configuration Variables ::
set connectionName="Wi-Fi"
set staticIP=192.168.1.107
set subnetMask=255.255.255.0
set defaultGateway=192.168.1.2

:: Change Nothing Below This Line ::
netsh interface ip set address %connectionName% static %staticIP% %subnetMask% %defaultGateway% 1
netsh interface ip set dns %connectionName% static %defaultGateway%