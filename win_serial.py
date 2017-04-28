import ctypes,subprocess


class EmuSerial(object):
	def __init__(self,emulated_port_settings,nullmodem_port_settings):
		# Get this value from the output of the install to close the connection.
		self.emu_port = emulated_port_settings
		self.null_port = nullmodem_port_settings
		self.native = ctypes.CDLL("./com0com.dll")
		if(self.create_emulated_pair() == False):
			self.wipe_ports()
			self.create_emulated_pair()
		
	def find_port_number(self,target):
		p = subprocess.Popen(["setupc.exe","list"],stdout=subprocess.PIPE)
		for line in p.stdout:
			if(target in line):
				line = line.strip()
				line = line.replace("CNCA","").replace("CNCB","")
				ls = line.split(" ")[0]
				return int(ls)
		return -1
	
	def remove_port(self,port_number):
		result = self.native.MainA("PythonWindowsSerialEmu","--silent remove %d" % port_number)
		if(result == 0):
			return True
		
		return False
		
	def __del__(self):
		self.wipe_ports()
		
	def wipe_ports(self):
		# Check if SRC or DEST are already bound.
		cport_number = self.find_port_number(self.emu_port["endpoint"])
		print("Current Port Number: %d" % cport_number)
		if(cport_number != -1):
			result = self.remove_port(cport_number)
			if(result == False):
				print("Error Removing Port %d" % cport_number)

		cport_number = self.find_port_number(self.null_port["endpoint"])
		print("Current Port Number: %d" % cport_number)
		if(cport_number != -1):
			result = self.remove_port(cport_number)
			if(result == False):
				print("Error Removing Port %d" % cport_number)
				
	def create_emulated_pair(self):
		self.wipe_ports()
				
		command_str = "--silent install"
		# Set up From Port Commands
		from_str = ["PortName=%s" % self.emu_port["endpoint"]]
		if(self.emu_port["cts_dsr"] == True):
			from_str.append("cts=rdtr")
		if(self.emu_port["strict_baud_emu"] == True):
			from_str.append("EmuBR=yes")
		from_str = ",".join(from_str)
		# Set up To Port Commands
		to_str = ["PortName=%s" % self.null_port["endpoint"]]
		if(self.null_port["cts_dsr"] == True):
			to_str.append("cts=rdtr")
		if(self.null_port["strict_baud_emu"] == True):
			to_str.append("EmuBR=yes")
		to_str = ",".join(to_str)		
		command_str = " ".join([command_str,from_str,to_str])

		result = self.native.MainA("PythonWindowsSerialEmu",command_str)
		if(result == 0):
			return True
		else:
			return False
		
		
if(__name__=="__main__"):
	print("Testing...")
	emulated_port_settings = {
	"endpoint":"COM4",
	"cts_dsr":False,
	"strict_baud_emu":True
	}
	
	nullmodem_port_settings = {
	"endpoint":"COM14",
	"cts_dsr":False,
	"strict_baud_emu":True
	}
	
	es = EmuSerial(emulated_port_settings,nullmodem_port_settings)
	del(es)
