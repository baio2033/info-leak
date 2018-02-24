'''
LNK file parser

'''

import sys, os
from struct import *
from datetime import datetime, timedelta, tzinfo
from calendar import timegm

EPOCH_AS_FILETIME = 116444736000000000  # January 1, 1970 as MS file time
HUNDREDS_OF_NANOSECONDS = 10000000

line = ""

def convertTime(t):
  return datetime.fromtimestamp(t)

def filetime_to_dt(ft):
  if ft == 0:
    return 0
  return str(datetime.utcfromtimestamp((ft - EPOCH_AS_FILETIME) / HUNDREDS_OF_NANOSECONDS))

def file_size_converter(size):
    magic = lambda x: str(round(size/round(x/1024), 2))
    size_in_int = [int(1 << 10), int(1 << 20), int(1 << 30), int(1 << 40), int(1 << 50)]
    size_in_text = ['B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB']
    for i in size_in_int:
        if size < i:
            g = size_in_int.index(i)
            position = int((1024 % i) / 1024 * g)
            ss = magic(i)
            return ss + ' ' + size_in_text[position]

def lnkflag_parser(flag):
  lst = []
  if flag & 2**0 == 2**0:
    lst.append("LNKTargetIDList")
  if flag & 2**1 == 2**1:
    lst.append("LNKInfo")
  if flag & 2**2 == 2**2:
    lst.append("Name")
  if flag & 2**3 == 2**3:
    lst.append("RelativeDir")
  if flag & 2**4 == 2**4:
    lst.append("WorkingDir")
  if flag & 2**7 == 2**7:
    lst.append("IsUnicode")
  return lst

def fileattr_parser(flag):
  lst = []
  if flag & 2**0 == 2**0:
    lst.append("FILE_ATTRIBUTE_READONLY")
  if flag & 2**1 == 2**1:
    lst.append("FILE_ATTRIBUTE_HIDDEN")
  if flag & 2**2 == 2**2:
    lst.append("FILE_ATTRIBUTE_SYSTEM")
  if flag & 2**4 == 2**4:
    lst.append("FILE_ATTRIBUTE_DIRECTORY")
  if flag & 2**5 == 2**5:
    lst.append("FILE_ATTRIBUTE_ARCHIVE")
  if flag & 2**7 == 2**7:
    lst.append("FILE_ATTRIBUTE_NORMAL")
  if flag & 2**8 == 2**8:
    lst.append("FILE_ATTRIBUTE_TEMPORARY")
  if flag & 2**9 == 2**9:
    lst.append("FILE_ATTRIBUTE_SPARSE_FILE")
  if flag & 2**10 == 2**10:
    lst.append("FILE_ATTRIBUTE_REPARSE_POINT")
  if flag & 2**11 == 2**11:
    lst.append("FILE_ATTRIBUTE_COMPRESSED")
  if flag & 2**12 == 2**12:
    lst.append("FILE_ATTRIBUTE_NOT_CONTENT_INDEXED")
  if flag & 2**13 == 2**13:
    lst.append("FILE_ATTRIBUTE_ENCRYPTED")
  return lst

def targetidlist_parser(f, offset):
  f.seek(offset)
  idlistSize = unpack("<H", f.read(2))[0]
  #print "\n[+] IDListSize : ", hex(idlistSize)
  while True:
    itemIDSize = unpack("<H", f.read(2))[0]
    if itemIDSize == 0:
      break
    tmp_str = str(itemIDSize-2) + "s"
    data = unpack(tmp_str, f.read(itemIDSize-2))
    #print "[+] Data : ", data
  return f.tell()

def lnkinfo_parser(f, offset, line):
  f.seek(offset)
  lnkinfoSize = unpack("<I", f.read(4))[0]
  #print "\n[+] LNKInfoSize: ", hex(lnkinfoSize)
  lnkinfoHeaderSize = unpack("<I", f.read(4))[0]
  #print "\n[+] lnkinfoHeaderSize: ", hex(lnkinfoHeaderSize)
  lnk_flag = unpack("<I", f.read(4))[0]

  volumeID_basepath = False
  if lnk_flag & 1 == 1:
    volumeID_basepath = True

  if volumeID_basepath:
    volumeID_offset = unpack("<I", f.read(4))[0]
    basepath_offset = unpack("<I", f.read(4))[0]

    f.seek(offset + volumeID_offset)
    VolumeIDSize = unpack("<I", f.read(4))[0]
    DriveType = unpack("<I", f.read(4))[0]
    DriveSerial = unpack("<I", f.read(4))[0]
    print "[+] DriveSerial: ", hex(DriveSerial)
    VolumeLabelOffset = unpack("<I", f.read(4))[0]
    f.seek(offset + VolumeLabelOffset)

    volumeID = ""
    while True:
      c = f.read(1)
      if c == "\x00":
        break
      volumeID += c

    print "[+] VolumeID: ", volumeID

    f.seek(offset + basepath_offset)
    basepath = ""
    while True:
      c = f.read(1)
      if c == "\x00":
        break
      basepath += c

    print "[+] BasePath: ", basepath

    line += basepath + "," + volumeID + "\n"
  return offset + lnkinfoSize, line

def shelllnk_parser(f):
  global line
  f.seek(0)
  shell_lnk = unpack_from('<IQQIIQQQIIIHHII',f.read(0x4c), 0)

  header_size = shell_lnk[0]
  lnk_flag = shell_lnk[3]
  flag_info = lnkflag_parser(lnk_flag)
  file_attribute = fileattr_parser(shell_lnk[4])
  fattr = ""
  for attr in file_attribute:
    fattr += "[" + attr + "]" 
  ctime = filetime_to_dt(shell_lnk[5])
  atime = filetime_to_dt(shell_lnk[6])
  wtime = filetime_to_dt(shell_lnk[7])
  file_size = shell_lnk[8]

  print "[+] Target Creation Time : ", ctime
  print "[+] Target Access Time : ", atime
  print "[+] Target Modified Time : ", wtime
  print "[+] File Size : ", file_size
  print "[+] File Attribution : ", file_attribute
  #print "[+] Flag : ", flag_info

  line += str(ctime) + "," + str(atime) + "," + str(wtime) + "," + str(file_size) + "," + fattr + ","

  offset = header_size
  if "LNKTargetIDList" in flag_info:
    offset = targetidlist_parser(f, offset)
  if "LNKInfo" in flag_info:
    offset, line = lnkinfo_parser(f, offset,line)

  #print line



if __name__ == "__main__":
  print('\nType Volume name to analyze LNK files...(ex. c:/)\n')
  vol = raw_input('>> ')

  print('\n>> Searching .LNK files\n')
  print('--------------------------------------------------------\n')

  if os.path.isdir('./result') == False:
    os.mkdir('./result')

  result = open('./result/lnk_result.csv','w')
  result.write('FileName,mtime,atime,ctime,Target-mtime,Target-atime,Target-ctime,FileSize,FileAttr,BasePath,VolumeID\n')

  cnt = 0
  for (path, dir, files) in os.walk(vol):
    if "AppData\\Roaming\\Microsoft\\Windows\\Recent" in path:
      for filename in files:
        ext = os.path.splitext(filename)[-1]
        if ext.lower() == '.lnk':
          global line
          line = ""
          fpath = path + "\\" + filename
          #print(fpath)
          mtime = str(convertTime(os.path.getmtime(fpath)))
          atime = str(convertTime(os.path.getatime(fpath)))
          ctime = str(convertTime(os.path.getctime(fpath)))

          line += filename + "," + mtime + "," + atime + "," + ctime + ","
          f = open(path + "\\" + filename, "rb")
          shelllnk_parser(f)
          f.close()
          result.write(line)
          cnt += 1

  result.close()
  print('-------------------------------------------------------\n')
  print('>> Finished (Total %d .lnk files detected)\n' % cnt)