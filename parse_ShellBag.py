import os, sys, pytsk3, re
from Registry import Registry
from BinaryParser import OverrunBufferException
from ShellItems import SHITEMLIST

shellbag_list = []

def cvtPath(p):
	p = p[2:].replace("\\","/")
	return p

def exportHIVE(vol):
	if os.path.isdir('./export') == False:
		os.mkdir('./export')

	img = pytsk3.Img_Info('\\\\.\\'+vol)
	fs_info = pytsk3.FS_Info(img)

	userprofile = os.path.expandvars("%userprofile%")
	
	ntuser_ = cvtPath(userprofile + "\\NTUSER.DAT")
	usrclass_ = cvtPath(userprofile + "\\AppData\\Local\\Microsoft\\Windows\\UsrClass.dat")
	exportList = []
	exportList.append(ntuser_)
	exportList.append(usrclass_)

	for p in exportList:
		f = fs_info.open(p)
		name = f.info.name.name.decode('utf-8')
		offset = 0
		if f.info.meta == None:
			continue
		size = f.info.meta.size
		BUFF_SIZE = 1024 * 1024
		data = open('./export/'+name,'wb')
		while offset < size:
			av_to_read = min(BUFF_SIZE, size - offset)
			d = f.read_random(offset,av_to_read)
			if not d: break
			data.write(d)
			offset += len(d)

	print "\nExport Completed\n"

def cvtDate(d):
	try:
		return d.strftime("%Y/%m/%d %H:%M:%S")
	except:
		return "1970/01/01 00:00:00"

def parse_shellbags(bagmru_key, bags_key, key, bag_pre, path_pre):
	try:
		slot = key.value("NodeSlot").value()
		for bag in bags_key.subkey(str(slot)).subkeys():
			for val in [val for val in bag.values() if "ItemPos" in val.name()]:
				buf = val.value()

				blk = SHITEMLIST(buf, 0, False)
				offset = 0x10

				while True:
					offset += 0x8
					size = block.unpack_word(offset)
					if size == 0:
						break
					elif size < 0x15:
						pass
					else:
						item = blk.get_item(offset)
						shellbag_list.append({
							"path": path_pre + "\\" + item.name(),
							"mtime": cvtDate(item.m_date()),
							"atime": cvtDate(item.a_date()),
							"crtime": cvtDate(item.cr_date()),
							"source": bag.path() + " @ " + hex(item.offset()),
							"regsource": bag.path() + "\\" + val.name(),
							"klwt": cvtDate(key.timestamp())
							})
					offset += size
	except Registry.RegistryValueNotFoundException:
		pass
	except Registry.RegistryKeyNotFoundException:
		print "[-] no key"
		pass
	except:
		print "[-] error"

	for val in [val for val in key.values() if re.match("\d+", val.name())]:
		path = ""
		try:
			lst = SHITEMLIST(val.value(), 0, False)
			for item in lst.items():
				path = path_pre + "\\" + item.name()
				shellbag_list.append({
					"path": path,
					"mtime": cvtDate(item.m_date()),
					"atime": cvtDate(item.a_date()),
					"crtime": cvtDate(item.cr_date()),
					"source": key.path() + " @ " + hex(item.offset()),
					"klwt": cvtDate(key.timestamp())
					})
		except OverrunBufferException:
			print key.path()
			print val.name()
			raise

		parse_shellbags(bagmru_key, bags_key, key.subkey(val.name()),bag_pre + "\\" + val.name(), path)

if __name__=="__main__":

	print "\nExtract Hive files\n"

	rootVol = os.environ['WINDIR'][:2]
	exportHIVE(rootVol)

	hive_file = ['./export/NTUSER.DAT','./export/UsrClass.dat']

	shell_path = ["Local Settings\\Software\\Microsoft\\Windows\\ShellNoRoam",
				  "Local Settings\\Software\\Microsoft\\Windows\\Shell"]

	result = []

	for hive in hive_file:
		reg = Registry.Registry(hive)

		for sp in shell_path:
			try:
				shell_key = reg.open(sp)
				bagmru_key = shell_key.subkey("BagMRU")
				bags_key = shell_key.subkey("Bags")
				parse_shellbags(bagmru_key, bags_key, bagmru_key, "", "")
			except Registry.RegistryKeyNotFoundException:
				pass
			except Exception:
				print "[-] Error during parsing ", sp

	if os.path.isdir('./result') == False:
		os.mkdir('./result')

	result = open('./result/shellbag.csv','w')
	line = "Path,RegTime,mtime,atime,ctime,Source\n"
	result.write(line)

	for t in shellbag_list:
		line = t['path'] + "," + t['klwt'] + "," + t['mtime'] + "," + t['atime'] + "," + t['crtime'] + "," + t['source'] + "\n"
		print line
		try:
			result.write(line)
		except:
			#line = str(t['path']) + "," + t['mtime'] + "," + t['atime'] + "," + t['crtime'] + "," + t['source'] + "," + t['klwt'] + "\n"
			line = t['path'] + "\n"
			print line
	print "\n[+] Completed!\n"

	result.close()
