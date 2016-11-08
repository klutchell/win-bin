# Get-AppxPackage *3dbuilder* | Remove-AppxPackage
# Get-AppxPackage *windowscommunicationsapps* | Remove-AppxPackage
# Get-AppxPackage *officehub* | Remove-AppxPackage
# Get-AppxPackage *skypeapp* | Remove-AppxPackage
# Get-AppxPackage *solitairecollection* | Remove-AppxPackage
# Get-AppxPackage *Messaging* | Remove-AppxPackage

Get-AppxPackage -AllUsers | Remove-AppxPackage