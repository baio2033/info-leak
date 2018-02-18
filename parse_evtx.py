import Evtx.Evtx as evtx
import Evtx.Views as e_views
import sys, time, os


def getEvtID(record):
	start_idx = record.find('EventID Qualifiers="">')
	end_idx = record.find('</EventID>')
	data_idx = start_idx + len('EventID Qualifiers="">')
	return record[data_idx:end_idx]

def getEvtTime(record):
	start_idx = record.find('TimeCreated SystemTime="')
	end_idx = record.find('</TimeCreated>') - 2
	data_idx = start_idx + len('TimeCreate SystemTime="') + 1
	now = record[data_idx:end_idx]
	idx = now.find(".")
	now = now[:idx]
	return now

def get_child(node, tag, ns="{http://schemas.microsoft.com/win/2004/08/events/event}"):
	return node.find("%s%s"%(ns,tag))

def getUSBlog(evtxfile):
	result = open('usb_eventlog.csv','w')
	result.write('EventID,Time,Event,Product,SerialNum,lifetime\n')
	vendor, product, revision, serialNum, lifetime = "", "", "", "", ""
	with evtx.Evtx(evtxfile) as log:
		for record in log.records():
			preRecord = record.xml()
			preTime = getEvtTime(preRecord)
			break
	count = 0
	print("\n\n[+] Processing...\n")
	with evtx.Evtx(evtxfile) as log:
		for record in log.records():
			xmlstr = record.xml()
			evtid = int(getEvtID(xmlstr))
			evtTime = getEvtTime(xmlstr)

			if preTime == evtTime:
				count += 1
			else:
				if count==25 and preTime!=evtTime:
					state = "Mounted"
					result.write(str(evtid)+","+evtTime+","+state+","+product+","+serialNum+","+lifetime+"\n")
				elif count==8 and preTime!=evtTime:
					state="Unmounted"
					result.write(str(evtid)+","+evtTime+","+state+","+product+","+serialNum+","+lifetime+"\n")
				count=1
			offset = xmlstr.find("DISK&")
			if offset > 0:
				tmpOffset = xmlstr.find("lifetime") - 3
				tmp = xmlstr[offset:tmpOffset]
				startOffset = tmp.find("VEN_") + 4
				endOffset = tmp.find("&PROD")
				vendor = tmp[startOffset:endOffset]
				startOffset = tmp.find("PROD_") + 5
				endOffset = tmp.find("&amp;REV")
				product = tmp[startOffset:endOffset]
				startOffset = tmp.find("#") + 1
				endOffset = startOffset + tmp[startOffset:].find("&amp;")
				serialNum = tmp[startOffset:endOffset]
			offset = xmlstr.find("lifetime")
			if offset > 0:
				tmpOffset = xmlstr.find("xmlns:auto")
				tmp = xmlstr[offset:tmpOffset]
				startOffset = tmp.find("=\"{") + 2
				endOffset = tmp.find("}") + 1
				lifetime = tmp[startOffset:endOffset]
			preRecord = record.xml()
			preTime = getEvtTime(preRecord)
	result.close()
	print("\n[+] Completed!\n")


if __name__=="__main__":
	evt_path = "C:\Windows\Sysnative\winevt\Logs\Microsoft-Windows-DriverFrameworks-UserMode%4Operational.evtx"
	evt_name = "Microsoft-Windows-DriverFrameworks-UserMode%4Operational.evtx"
	getUSBlog(evt_name)

