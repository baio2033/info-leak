import winreg

def subkeys(key):
    i = 0
    while True:
        try:
            subkey = winreg.EnumKey(key, i)
            yield subkey
            i += 1
        except WindowsError:
            break

def traverse_registry_tree(hkey, keypath, tabs=0):
    key = winreg.OpenKey(hkey, keypath, 0, winreg.KEY_READ)
    for subkeyname in subkeys(key):
        if "{53f56307-b6bf-11d0-94f2-00a0c91efb8b}" in keypath:
            parseInfo(str(subkeyname))
        subkeypath = "%s\\%s" % (keypath, subkeyname)
        if "{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}" in keypath:
            #print(subkeyname)
            if "Properties" in subkeyname:
                continue 
            if "Volume" in subkeyname:
                sub = winreg.OpenKey(hkey, keypath + "\\" + subkeyname, 0, winreg.KEY_READ)
                try:
                    i = 0
                    while True:
                        ret = winreg.EnumValue(sub,i)
                        if ret[0] == "DeviceInstance":
                            print("DeviceInstance : ", ret[1])
                        i += 1
                except WindowsError:
                    pass
                       
        traverse_registry_tree(hkey, subkeypath, tabs+1)

def parseInfo(val):
    if val.find("&Ven_") > 0 and val.find("&Prod_") > 0 and val.find("#Disk") > 0:
        startOff = val.find("&Ven_")
        startOff += len("&Ven_")
        endOff = startOff + val[startOff:].find("&Prod_")
        vendor = val[startOff:endOff]
        startOff = val.find("&Prod_")
        startOff += len("&Prod_")
        endOff = startOff + val[startOff:].find("#")
        product = val[startOff:endOff]
        print("#######################################")
        print("Vendor : ", vendor)
        print("Product : ", product) 

if __name__=="__main__":

    print("\nDevice GUID\n")
    keypath = r"SYSTEM\\ControlSet001\\Control\\DeviceClasses\\{53f56307-b6bf-11d0-94f2-00a0c91efb8b}"
    traverse_registry_tree(winreg.HKEY_LOCAL_MACHINE, keypath)
    print("\n--------------------------------------------------------\n")
    print("\nVolume GUID\n")
    keypath = r"SYSTEM\\ControlSet001\\Control\\DeviceClasses\\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}"
    traverse_registry_tree(winreg.HKEY_LOCAL_MACHINE, keypath)