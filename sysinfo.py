import wmi
ws = wmi.WMI(namespace='root/Microsoft/Windows/Storage')
for d in ws.MSFT_PhysicalDisk():
    print(d.BusType, d.MediaType, d.Model, d.SlotNumber)
