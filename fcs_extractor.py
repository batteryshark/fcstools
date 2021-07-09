'''
	FileCabinet Solution (FCS) Extraction Utility 
		2014 - rFx
'''

import os,sys,struct,zlib,io,shutil,subprocess
OP_RT = os.getcwd()
DRAWER_MAP = {}
DRAWERDB_FILENAME = "CltList.cc$"
OUT_BS = "E:\\OUT\\"
OUT_BASE = "E:\\OUT\\"
FILENAME_BLACKLIST = ['"','*','/',':','<','>','?','\\','|']
FOLDER_BLACKLIST = ['"','*',':','<','>','?','|','&']


DOC_FORMAT = {
"WordDocument":"doc",
'Workbook':"xls",
'PowerPoint Document':"ppt"
}
FILENAME_BLACKLIST = ['"','*','/',':','<','>','?','\\','|']
FOLDER_BLACKLIST = ['"','*',':','<','>','?','|','&']

	
def proc_olegroup(in_path):
	os.chdir(in_path)
	out_path = OUT_BASE
	f = open("INFO_OLEITEM","rb")
	crap20 = f.read(20)
	tyskip_sz = struct.unpack("<H",f.read(2))[0]
	tyskip = f.read(tyskip_sz)
	dirname_sz =  struct.unpack("<H",f.read(2))[0]
	dirname = f.read(dirname_sz)
	crap4 = f.read(4)
	owner_sz = struct.unpack("<H",f.read(2))[0]
	owner = f.read(owner_sz)
	crap22 = f.read(22)
	pdir_sz = struct.unpack("<H",f.read(2))[0]
	pdir = f.read(pdir_sz)
	for ele in FILENAME_BLACKLIST:
		dirname = dirname.replace(ele,"")
	if(pdir_sz < 2 or pdir_sz > 255):
		pdir = ""
		out_path = os.path.join(OUT_BASE)
	else:
		for ele in FOLDER_BLACKLIST:
			pdir = pdir.replace(ele,"")
		out_path = os.path.join(OUT_BASE,pdir)
	f.close()
	
	if(not os.path.exists(out_path)):
		os.makedirs(out_path)
	
	for subdir, dirs, files in os.walk("."):
		for dir in dirs:
			for d in DOC_FORMAT.keys():
				if(os.path.exists(os.path.join(dir,d))):
					dirname = "%s.%s" % (dirname,DOC_FORMAT[d])
			
				if(len(os.path.join(out_path,dirname)) > 260):
		
					if("Page" in dirname):
						dirname = dirname[dirname.find("Page"):]
						out_path = os.path.join(OUT_BASE,"%s" % dirname)
					else:
						if(len(OUT_BASE) > 260):
							print("PATH > 260")
							exit(1)
						out_path = os.path.join(OUT_BASE,"tbg%d" % tbg_cnt)
						tbg_cnt += 1
						print("WARNING: Old Path > 260 and can't find Page")
			for subdir0, d0, files0 in os.walk(dir):
				for d in d0:
					shutil.rmtree(os.path.join(dir,d))
				break
			subprocess.call([os.path.join(OP_RT,'repack_ole.exe'),'tmp','%s' % dir])
			shutil.copyfile("tmp",os.path.join(out_path,dirname))
	os.chdir("..")
	
def proc_imgrp(in_path):
	os.chdir(in_path)
	try:
		f = open("INFO_IMGRP","rb")
	except:
		os.chdir("..")
		return
	info_id = f.read(4)
	dirname_sz = struct.unpack("<H",f.read(2))[0]
	dirname = f.read(dirname_sz)
	num_pages = struct.unpack("<I",f.read(4))[0]
	
	crap_initial = f.read(4 * num_pages)
	crap_20 = f.read(0x14)
	pdir_sz = struct.unpack("<H",f.read(2))[0]

	if(pdir_sz < 2 or pdir_sz > 255):
		pdir = ""

		for ele in FOLDER_BLACKLIST:
			dirname = dirname.replace(ele,"")
		outpath = os.path.join(OUT_BASE,dirname)
	else:
		pdir = f.read(pdir_sz)
		for ele in FOLDER_BLACKLIST:
			dirname = dirname.replace(ele,"")
			pdir = pdir.replace(ele,"")
		outpath = os.path.join(OUT_BASE,pdir,dirname)
	f.close()
	try:	
		if(not os.path.exists(outpath)):
			os.makedirs(outpath)
		outpath = outpath.replace("/","")
		for subdir, dirs, files in os.walk("."):
			for dir in dirs:
				proc_imgrpfile(dir,outpath,dirname)
	except:
		with open(os.path.join(OP_RT,"log.txt"), "w") as text_file:
			text_file.write("%s %s Failed to Extract\n" % (in_path,outpath))
	os.chdir("..")


def proc_imgrpfile(in_path,out_path,mask_name):
	os.chdir(in_path)
	tbg_cnt = 1
	out_path = out_path.rstrip()
	old_out_path = out_path

	if(os.path.exists("INFO_IMG")):
		f = open("INFO_IMG","rb")
		info_id = f.read(4)
		filename_sz = struct.unpack("<H",f.read(2))[0]
		filename = f.read(filename_sz)
		for ele in FILENAME_BLACKLIST:
			filename = filename.replace(ele,"")
		extension_sz = struct.unpack("<H",f.read(2))[0]
		extension = f.read(extension_sz)
		f.close()
		if(not os.path.exists(out_path)):
			os.makedirs(out_path)
		out_path = os.path.join(out_path,"%s.%s" % (filename,extension))
	else:
		if(not os.path.exists(out_path)):
			os.makedirs(out_path)
		filename = "UNTITLED"
		extension = "bin"
		out_path = os.path.join(out_path,"UNTITLED.bin")
	if(len(out_path) + len("%s.%s" % (filename,extension)) > 260):
		
		if("Page" in filename):
			filename = filename[filename.find("Page"):]
			out_path = os.path.join(old_out_path,"%s.%s" % (filename,extension))
		else:
			if(len(old_out_path) > 260):
				print("PATH > 260")
				exit(1)
			out_path = os.path.join(old_out_path,"%s.%s" % ("Page 1-%d" % tbg_cnt,extension))
			tbg_cnt += 1
			print("WARNING: Old Path > 260 and can't find Page")
	
	shutil.copyfile("DATA_IMG",out_path)

			
	os.chdir("..")	
	


def proc_pgrpfile(in_path,out_path,mask_name):
	os.chdir(in_path)
	f = open("INFO_PG","rb")
	info_id = f.read(4)
	filename_sz = struct.unpack("<H",f.read(2))[0]
	filename = f.read(filename_sz)
	f.close()
	
	#Because these files are zlib compressed emf files...
	f = open("DATA_PG","rb")
	#Skip the header.
	f.seek(8)
	data = zlib.decompress(f.read())
	f.close()
	out_path = out_path.rstrip()
	for ele in FILENAME_BLACKLIST:
		filename = filename.replace(ele,"")
	if(out_path.endswith(".pdf")):
		filename = "%s.pdf" % filename
		out_path = out_path[:-4]
		out_path = out_path.rstrip()
	if(out_path.endswith(".xls")):
		filename = "%s.xls" % filename
		out_path = out_path[:-4]
		out_path = out_path.rstrip()
	else:
		filename = "%s.EMF" % filename
	if(len(os.path.join(out_path, filename)) > 260):
		if("Page" in filename):
			filename = filename[filename.find("Page"):]
			
		else:
			if(len(out_path) > 260):
				print("PATH > 260")
				exit(1)
			filename = "tbg_%d.EMF" % tbg_cnt
			tbg_cnt += 1
			print("WARNING: Old Path > 260 and can't find Page")
	
	if(not os.path.exists(out_path)):
		os.makedirs(out_path)
	try:
		f = open(os.path.join(out_path, filename),"wb")
		f.write(data)
		f.close()
	except:
		with open(os.path.join(OP_RT,"log.txt"), "w") as text_file:
			text_file.write("%s Failed to Extract\n" % os.path.join(out_path, filename))
		
	os.chdir("..")

def proc_pgrp(in_path):
	os.chdir(in_path)
	try:
		f = open("INFO_PGRP","rb")
	except:
		os.chdir("..")
		return
	info_id = f.read(4)
	dirname_sz = struct.unpack("<H",f.read(2))[0]
	dirname = f.read(dirname_sz)
	crap_20 = f.read(0x14)
	pdir_sz = struct.unpack("<H",f.read(2))[0]
	for ele in FOLDER_BLACKLIST:
		dirname = dirname.replace(ele,"")	
	if(pdir_sz < 2 or pdir_sz > 255):
		pdir = ""
		outpath = os.path.join(OUT_BASE,dirname)
	else:
		pdir = f.read(pdir_sz)
		for ele in FOLDER_BLACKLIST:
			pdir = pdir.replace(ele,"")
		outpath = os.path.join(OUT_BASE,pdir,dirname)
	f.close()
	if(outpath.endswith(".xls") or outpath.endswith(".pdf")):
		outpath = outpath[:-4]
	if(outpath.endswith("...")):
		outpath = outpath[:-3]
	if(not os.path.exists(outpath)):
		os.makedirs(outpath)
	
	for subdir, dirs, files in os.walk("."):
		for dir in dirs:
			proc_pgrpfile(dir,outpath,dirname)
	os.chdir("..")

def proc_image(in_path):
	os.chdir(in_path)
	f = open("INFO_IMG","rb")
	info_id = f.read(4)
	filename_sz = struct.unpack("<H",f.read(2))[0]
	filename = f.read(filename_sz)
	for ele in FILENAME_BLACKLIST:
		filename = filename.replace(ele,"")
	extension_sz = struct.unpack("<H",f.read(2))[0]
	extension = f.read(extension_sz)
	crap_8 = f.read(8)
	owner_sz = struct.unpack("<H",f.read(2))[0]
	owner = f.read(owner_sz)
	crap_20 = f.read(20)
	pdir_sz = struct.unpack("<H",f.read(2))[0]
	pdir = f.read(pdir_sz)
	f.close()
	if(not os.path.exists(os.path.join(OUT_BASE,pdir))):
		os.makedirs(os.path.join(OUT_BASE,pdir))
	out_path = os.path.join(OUT_BASE,pdir,"%s.%s" % (filename,extension))
	
	shutil.copyfile("DATA_IMG",out_path)
	os.chdir("..")
	

def proc_pdf(in_path):
	#PDF CODE
	os.chdir(in_path)
	if(not os.path.exists(OUT_BASE)):
		os.makedirs(OUT_BASE)
	try:
		f = open("INFO_PDF","rb")
	except:
		os.chdir("..")
		return
	info_id = f.read(4)
	filename_sz = struct.unpack("<H",f.read(2))[0]
	filename = f.read(filename_sz)
	for ele in FILENAME_BLACKLIST:
		filename = filename.replace(ele,"")
	hash = f.read(4)
	owner_sz = struct.unpack("<H",f.read(2))[0]
	owner = f.read(owner_sz)
	opath_sz = struct.unpack("<H",f.read(2))[0]
	opath = f.read(opath_sz)
	crap = f.read(0x1C)
	pdir_sz = struct.unpack("<H",f.read(2))[0]
	if(pdir_sz > 1):
		pdir = f.read(pdir_sz)
		out_path = os.path.join(OUT_BASE,pdir,"%s.pdf" % filename)
		if(not os.path.exists(os.path.join(OUT_BASE,pdir))):
			os.makedirs(os.path.join(OUT_BASE,pdir))
	else:
		out_path = os.path.join(OUT_BASE,"%s.pdf" % filename)
		if(not os.path.exists(OUT_BASE)):
			os.makedirs(OUT_BASE)
		
	f.close()

	
	try:
		shutil.copyfile("PDFDATA",out_path)
	except:
		pass
	os.chdir("..")


def proc_drawer(in_path,out_path):
	global OUT_BASE
	os.chdir(in_path)
	if(len(in_path) > 8):
		in_fulldrawerid = in_path.replace(".","")
	else:
		in_fulldrawerid = in_path
	out_path = os.path.join(out_path,"%s (%s)" % (DRAWER_MAP[in_path],in_fulldrawerid))
	
	OUT_BASE = out_path
	
	if(not os.path.exists(out_path)):
		os.makedirs(out_path)
	else:
		pass
		#os.system("del /F /Q %s" % out_path)
		#print(out_path)
		#exit(1)
		#os.makedirs(out_path)
		
	if(not os.path.exists("tmp")):
		os.makedirs("tmp")
	for root, dirs, files in os.walk("."):
		for file in files:
			if(file.endswith("cc$")):
				continue
			if(not file.startswith("CHAMP")):
				subprocess.call([os.path.join(OP_RT,'7z.exe'),'x','-otmp','-y',file])
		break
	os.chdir("tmp")
	for subdir, dirs, files in os.walk("."):
		for dir in dirs:
			print(dir)
			if(dir.startswith("PDF")):
				proc_pdf(dir)
			if(dir.startswith("IMAGE")):
				proc_image(dir)
			if(dir.startswith("PGRP")):
				proc_pgrp(dir)
			if(dir.startswith("IGRP")):
				proc_imgrp(dir)
			if(dir.startswith("OLE")):
				proc_olegroup(dir)
		break	
		
	os.chdir("..")
	os.system("del /F /Q tmp")
	os.chdir("..")

def get_drawermaps(in_path):
	f = open(os.path.join(in_path,DRAWERDB_FILENAME),"rb")
	ddbc = f.read()

	f.close()
	for subdir, dirs, files in os.walk(in_path):
		for dir in dirs:

			#Skip sysdata
			if(dir == "$sysdata"):
				continue
			ddb = ddbc
			dirfull = dir.replace(".","")
			#Hack to ensure we have the right drawer.
			cci_offset = ddb.rfind(dirfull)
			print(dirfull)
			ddb = ddb[cci_offset:]
			print(cci_offset)
			#CCI has a max filename size of 8 (DOS FORMAT 8.3)
			if(len(dir) > 8):
				sd = dir[:8]
			else:
				sd = dir
			cci_offset = ddb.find("%s.cci" % sd)
			if(cci_offset == -1):
				with open(os.path.join(OP_RT,"log.txt"), "w") as text_file:
					text_file.write("Drawer %s is not in DB\n" % (dir))
				DRAWER_MAP[dir] = dirfull
				continue
			cci_offset += len("%s.cci" % sd)
			dname_sz = struct.unpack("<H",ddb[cci_offset:cci_offset+2])[0]
			dname = ddb[cci_offset+2:cci_offset+2+dname_sz]
			if(dname == "" or dname == " "):
				dname = sd
			DRAWER_MAP[dir] = dname
			
		break
		exit(1)
	
	
	
def proc_cabinet(in_path):
	get_drawermaps(in_path)
	global OUT_BASE
	os.chdir(in_path)
	out_path = os.path.join(OUT_BASE,in_path)
	if(not os.path.exists(out_path)):
		os.makedirs(out_path)
		
	for subdir, dirs, files in os.walk("."):
		for dir in dirs:
			if(dir.startswith("$")):
				continue
			print("Processing Drawer: %s" % dir)
			proc_drawer(dir,out_path)
		break
	OUT_BASE = OUT_BS
	os.chdir("..")

if(__name__=="__main__"):
	for subdir, dirs, files in os.walk("."):
		for dir in dirs:
			if("#TOFIX_" in dir):
				continue
			print("Processing Cabinet: %s" % dir)
			proc_cabinet(dir)
		break
		
	print("Done!")