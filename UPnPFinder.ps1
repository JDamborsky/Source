$UPnPFinder = New-Object -ComObject UPnP.UPnPDeviceFinder
$UPnPFinder.FindByType("upnp:rootdevice", 0) | 
  Select-Object ModelName, FriendlyName, PresentationUrl |
  Sort-Object ModelName