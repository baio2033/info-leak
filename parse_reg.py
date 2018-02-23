import _winreg, os
from datetime import datetime

EPOCH_AS_FILETIME = 116444736000000000  # January 1, 1970 as MS file time
HUNDREDS_OF_NANOSECONDS = 10000000

USB_List = []

def filetime_to_dt(ft):
  if ft == 0:
    return 0
  return str(datetime.utcfromtimestamp((ft - EPOCH_AS_FILETIME) / HUNDREDS_OF_NANOSECONDS))

def keyInfo(key):
    ret = _winreg.QueryInfoKey(key)
    return ret

def subkeys(key):
    i = 0
    while True:
        try:
            subkey = _winreg.EnumKey(key, i)
            yield subkey
            i += 1
        except WindowsError:
            break

def traverse_registry_tree(hkey, keypath, tabs=0):
    key = _winreg.OpenKey(hkey, keypath, 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
    if "MountedDevices" in keypath:
        try:
            i = 0
            while True:
                ret = _winreg.EnumValue(key,i)
                print "#######################################"
                try:
                    tmp = str(ret[1].decode('utf-16'))
                    print ret[0], ":", tmp
                    if "DosDevices" in ret[0]:
                        parseLabel(ret[0],tmp)
                except:
                    print(ret[0], ":", ret[1].decode('utf-16'))
                print "LastModifiedTime : ", filetime_to_dt(keyInfo(key)[2])
                i += 1
        except WindowsError:
            pass
        return
    else:
        for subkeyname in subkeys(key):
            #print subkeyname)
            subkeypath = "%s\\%s" % (keypath, subkeyname)
            if "a5dcbf10" in keypath:
                #print subkeypath[:-2])
                sub = _winreg.OpenKey(hkey, subkeypath[:-2], 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
                try:
                    i = 0
                    while True:
                        ret = _winreg.EnumValue(sub,i)
                        if ret[0] == "DeviceInstance":
                            print "#######################################"
                            print "DeviceInstance : ", ret[1]
                            parseSer(ret[1], key)
                        i += 1
                except WindowsError:
                    pass
            elif "53f5630d" in keypath:
                if "Properties" == subkeyname:
                    continue 
                if "Volume" in subkeyname:
                    sub = _winreg.OpenKey(hkey, subkeypath, 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
                    try:
                        i = 0
                        while True:
                            ret = _winreg.EnumValue(sub,i)
                            if ret[0] == "DeviceInstance":
                                print "#######################################"
                                print "DeviceInstance : ", ret[1]
                                print "LastModifiedTime : ", filetime_to_dt(keyInfo(sub)[2])
                                parseInfo(subkeyname,sub)
                            i += 1
                    except WindowsError:
                        pass
            elif "53f56307" in keypath:
                sub = _winreg.OpenKey(hkey, subkeypath, 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
                parseInfo(subkeyname,sub)
            elif "Windows Portable Devices" in keypath:
                sub = _winreg.OpenKey(hkey, subkeypath, 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
                try:
                    i = 0
                    while True:
                        ret = _winreg.EnumValue(sub,i)
                        print "#######################################"
                        print ret[0], ":", ret[1]
                        parseFN(subkeyname,ret[1],sub)
                        i += 1
                except WindowsError:
                    pass
            elif "WpdBusEnumRoot\\UMB" in keypath:
                for t in USB_List:
                    if t['serial'] in subkeyname:
                        sub = _winreg.OpenKey(hkey, subkeypath, 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
                        regTime = filetime_to_dt(keyInfo(sub)[2])
                        t['LC'] = regTime
                        try:
                            i = 0
                            while True:
                                ret = _winreg.EnumValue(sub,i)
                                print "####################################"
                                print ret[0], ":", ret[1]
                                i += 1
                        except WindowsError:
                            pass
            elif "ProfileList" in keypath:
                print "#######################################"
                print "SID : ", subkeyname
                sub = _winreg.OpenKey(hkey, subkeypath, 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
                try:
                    i = 0
                    while True:
                        ret = _winreg.EnumValue(sub,i)
                        if ret[0] == "ProfileImagePath":
                            print "ProfileImagePath : ", ret[1]
                            print "LastModifiedTime : ", filetime_to_dt(keyInfo(key)[2])
                        i += 1
                except WindowsError:
                    pass
            traverse_registry_tree(hkey, subkeypath, tabs+1)

def parseLabel(header, val):
    l = header.split('\\')[-1]
    if val.find("&Ven_") != -1 and val.find("&Prod_") != -1 and val.find("Disk") != -1:
        startOff = val.find("&Ven_")
        startOff += len("&Ven_")
        endOff = startOff + val[startOff:].find("&Prod_")
        vendor = val[startOff:endOff]
        startOff = val.find("&Prod_")
        startOff += len("&Prod_")
        if val.find("&Rev_") != -1:
            endOff = startOff + val[startOff:].find("&Rev_")
            product = val[startOff:endOff]
            startOff = val.find("&Rev_")
            startOff += len("&Rev_")
            endOff = startOff + val[startOff:].find("#")
            version = val[startOff:endOff]
        else:
            endOff = startOff + val[startOff:].find("#")
            product = val[startOff:endOff]
            version = ""
        startOff = endOff + 1
        endOff = val[startOff:].find("#") + startOff
        serial = val[startOff:endOff]
        for t in USB_List:
            if serial == t['serial']:
                if l not in t['Label']:
                    t['Label'].append(l)

def parseInfo(val,key):
    if val.find("&Ven_") != -1 and val.find("&Prod_") != -1 and val.find("Disk") != -1:
        startOff = val.find("&Ven_")
        startOff += len("&Ven_")
        endOff = startOff + val[startOff:].find("&Prod_")
        vendor = val[startOff:endOff]
        startOff = val.find("&Prod_")
        startOff += len("&Prod_")
        if val.find("&Rev_") != -1:
            endOff = startOff + val[startOff:].find("&Rev_")
            product = val[startOff:endOff]
            startOff = val.find("&Rev_")
            startOff += len("&Rev_")
            endOff = startOff + val[startOff:].find("#")
            version = val[startOff:endOff]
        else:
            endOff = startOff + val[startOff:].find("#")
            product = val[startOff:endOff]
            version = ""
        startOff = endOff + 1
        endOff = val[startOff:].find("#") + startOff
        serial = val[startOff:endOff]
        regTime = filetime_to_dt(keyInfo(key)[2])
        print "#######################################"
        print "Vendor : ", vendor
        print "Product : ", product
        print "Version : ", version 
        print "Serial : ", serial
        print "First Connection Time after Booting : ", regTime
        print 

        tmp = {'vendor':vendor,'product':product,'version':version,'serial':serial,'FC':'','FCB':[regTime], 'LC':'', 'FriendlyName':'', 'Label':[]}
        flag = True
        for t in USB_List:
            if t == tmp:
                flag = False
                if t['FCB'] != tmp['FCB']:
                    for regt in tmp['FCB']:
                        t['FCB'].append(regt)
        if flag:
            USB_List.append(tmp)

def parseSer(val,key):
    startOff = val.find('&PID_')
    endOff = startOff + val[startOff:].find('\\')
    startOff = endOff + 1
    serial = val[startOff:]
    regTime = filetime_to_dt(keyInfo(key)[2])
    for t in USB_List:
        if serial in t['serial']:
            if regTime not in t['FCB']:
                t['FCB'].append(regTime)
                break
    print "First Connection : ", regTime

def parseSETUP(setuppath):
    sup = open(setuppath,'r')
    while True:
        line = sup.readline()
        if line:
            if "Device Install (Hardware initiated)" in line:
                tmp = line.split("\\")[-1]
                if "}#" in tmp:
                    tmp = tmp.split("}#")[-1]
                serial = tmp[:-2]
                for t in USB_List:
                    if serial in t['serial']:
                        tmptime = sup.readline()
                        tmptime = tmptime.split("start ")[-1][:-2]
                        if t['FC'] == "":
                            t['FC'] = tmptime
                            print "#######################################"
                            print "Serial : ", serial
                            print "Time : ", tmptime
        else:
            break

def parseFN(val,fn,key):
    regTime = filetime_to_dt(keyInfo(key)[2])
    if val.find("&VEN_") != -1 and val.find("&PROD_") != -1 and val.find("DISK") != -1:
        startOff = val.find("&VEN_")
        startOff += len("&VEN_")
        endOff = startOff + val[startOff:].find("&PROD_")
        vendor = val[startOff:endOff]
        startOff = val.find("&PROD_")
        startOff += len("&PROD_")
        if val.find("&REV_") != -1:
            endOff = startOff + val[startOff:].find("&REV_")
            product = val[startOff:endOff]
            startOff = val.find("&REV_")
            startOff += len("&REV_")
            endOff = startOff + val[startOff:].find("#")
            version = val[startOff:endOff]
        else:
            endOff = startOff + val[startOff:].find("#")
            product = val[startOff:endOff]
            version = ""
        startOff = endOff + 1
        endOff = val[startOff:].find("#") + startOff
        serial = val[startOff:endOff]
        print serial
        for t in USB_List:
            if serial in t['serial']:
                if t['FriendlyName'] == "":
                    t['FriendlyName'] = fn
                if regTime not in t['FCB']: 
                    t['FCB'].append(regTime)

        print "LastModifiedTime : ", regTime

if __name__=="__main__":

    if os.path.isdir('./result') == False:
        os.mkdir('./result')
    
    print "\n>> Device Class ID\n"
    keypath = r"SYSTEM\\ControlSet001\\Control\\DeviceClasses\\{53f56307-b6bf-11d0-94f2-00a0c91efb8b}"
    traverse_registry_tree(_winreg.HKEY_LOCAL_MACHINE, keypath)
    print "\n--------------------------------------------------------\n"

    print "\n>> Volume GUID\n"
    keypath = r"SYSTEM\\ControlSet001\\Control\\DeviceClasses\\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}"
    traverse_registry_tree(_winreg.HKEY_LOCAL_MACHINE, keypath)
    print "\n--------------------------------------------------------\n"

    print "\n>> USB GUID\n"
    keypath = r"SYSTEM\\ControlSet001\\Control\\DeviceClasses\\{A5DCBF10-6530-11D2-901F-00C04FB951ED}"
    traverse_registry_tree(_winreg.HKEY_LOCAL_MACHINE, keypath)
    print "\n--------------------------------------------------------\n"

    print "\n>> Portable devices\n"
    keypath = r"SOFTWARE\\Microsoft\\Windows Portable Devices\\Devices"
    traverse_registry_tree(_winreg.HKEY_LOCAL_MACHINE, keypath)
    print "\n--------------------------------------------------------\n"

    print "\n>> Mounted devices\n"
    keypath = r"SYSTEM\\MountedDevices"
    traverse_registry_tree(_winreg.HKEY_LOCAL_MACHINE, keypath)
    print "\n--------------------------------------------------------\n"

    print "\n>> WpdBusEnumRoot\\UMB\n"
    try:
        keypath = r"SYSTEM\\ControlSet001\\Enum\\WpdBusEnumRoot\\UMB"
        traverse_registry_tree(_winreg.HKEY_LOCAL_MACHINE, keypath)
    except:
        print "None"
    print "\n--------------------------------------------------------\n"

    
    print "\n>> ProfileList\n"
    keypath = r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList"
    traverse_registry_tree(_winreg.HKEY_LOCAL_MACHINE, keypath)
    print "\n--------------------------------------------------------\n"

    print "\n>> Setup.dev.api\n"
    setuppath = "C:\\Windows\\INF\\setupapi.dev.log"
    parseSETUP(setuppath)
    print "\n--------------------------------------------------------\n"

    result = open('./result/usb_registry.csv','w')
    result.write('Vendor,Product,Version,Serial,FirstConn(After Booting),FristConn,LastConn,FriendlyName,Label\n')
    print "\n>> Result\n"
    for t in USB_List:
        line = ""
        fcb = ""
        for fc in t['FCB']:
            fcb += fc + " / "
        lb = ""
        for l in t['Label']:
            lb += l + " / "
        line += t['vendor'] + "," + t['product'] + "," + t['version'] + "," + t['serial'] + "," + fcb + "," + t['FC'] + "," + t['LC'] + "," + t['FriendlyName'] + "," + lb + "\n"
        print line  
        result.write(line)

    result.close()

