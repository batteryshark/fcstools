#!/usr/bin/env python3
'''
License Generator for FileCabinet Solution
'''

import os,sys,struct


LIC_KEY = [0xAD, 0xDE, 0xEF, 0x00, 0xAB, 0xBE, 0xFC]
COMPANY = "Company Name".encode('ascii')
NAME = "Your Name".encode('ascii')
ADDRESS = "Address Line 1".encode('ascii')
CITY = "City".encode('ascii')
STATE = "NY".encode('ascii')
ZIP = "10108".encode('ascii')
PHONE = "210-555-5555".encode('ascii')

ZIP_4 = "0000".encode('ascii')
LIC_EXP = "121629".encode('ascii')
LIC_VER = "64608".encode('ascii')
VER = "120000".encode('ascii')
VER2 = "900".encode('ascii')
VER3 = "0".encode('ascii')
VER4 = "45".encode('ascii')

lic_data = b""

def gen_data_p2():
	p2 = b""
	p2 += ZIP_4
	p2 += LIC_EXP
	p2 += b"\x00" * 4
	p2 += b" "
	p2 += LIC_VER
	p2 += b"\x00"
	p2 += b" "
	p2 += VER
	p2 += b"\x20" * (16 - len(VER))
	p2 += VER2
	p2 +=  b"\x20" * (0x14 - len(VER2))
	p2 += VER3
	p2 += b"\x20" * (0x5 - len(VER3))
	p2 += VER4
	p2 +=  b"\x20" * (0x27 - len(VER2))	
	return p2
	

def gen_data_p1():
	p1 = b""
	p1 += COMPANY
	p1 += b" " * (0x28 - len(COMPANY))
	p1 += NAME
	p1 += b" " * (0x1E - len(NAME))
	p1 += ADDRESS
	p1 += b" " * (0x1E - len(ADDRESS))
	p1 += CITY
	p1 += b" " * (0x14 - len(CITY))
	p1 += STATE
	p1 += ZIP
	p1 += b"\xFF" * 5 #BECAUSE YOLO
	p1 += PHONE
	p1 += b" " * (0xF - len(PHONE))
	p1 += b"\x17\x0D"
	return p1

def encrypt_p2(data):
	data = bytearray(data)
	for i in range(0,len(data)):
		data[i] ^= LIC_KEY[i % len(LIC_KEY)]
	return data
	
def encrypt(data):
	outdata = bytearray()
	data = bytearray(data)
	chaff_val = 0
	i = 0
	k = 0
	while(i< len(data)-2):
		if(data[i] == 0x20):
			chaff_val +=1
		else:
			outdata+= struct.pack("B",(data[i] ^ LIC_KEY[k % len(LIC_KEY)]) % 256)
			k+=1
		i+=1
	
	outdata += b"\x00" * (chaff_val - 2)
	return outdata
	
	
data_p1 = gen_data_p1()
lic_data+=data_p1
lic_data+=encrypt(data_p1)
data_p2 = gen_data_p2()
lic_data+=data_p2
lic_data+=encrypt(data_p2)

lic_data+=b"\x00\x00\x00\x00NEW!          \xFF\xFF"

f = open("ZFCNAME.DAT","wb")
f.write(lic_data)
f.close()