import serial

# get user-data (CSV) as list-object
user_data = [150, 1.2, 0.05, 200, 300, 0.5, 0.5, 300]

# PACEMAKER PARAMETERS
KEY = 0xAA
LOWER_RATE_LIMIT = 0
ATRIAL_PULSE_WIDTH = 1
VENTRICULAR_PULSE_WIDTH = 2
VRP = 3
ARP = 4
ATRIAL_AMPLITUDE = 5
VENTRICULAR_AMPLITUDE = 6
AV_DELAY = 7
MESSAGE_LENGTH = 13
STATE = 0x00

# PACEMAKER SERIAL PARAMETERS
BAUDRATE = 115200
TIMEOUT = 5


''' PARAMETER PROCESSING '''

# lower rate lim
def LRL(num):
    return num

# atrial pulse width
def APW(num):
    return int(num*100)

# ventricular pulse width
def VPW(num):
    return int(num*100)
    
# VRP (split into two summands)
def fVRP(num):
    remainder = 1 if (num % 2 == 1) else 0
    return [num//2, num//2 + remainder]

# ARP
def fARP(num):
    remainder = 1 if (num % 2 == 1) else 0
    return [num//2, num//2 + remainder]

# atrial amplitude
def AA(num):
    return int(num*10)

# ventricular amplitude
def VA(num):
    return int(num*10)

# AV delay
def AVD(num):
    remainder = 1 if (num % 2 == 1) else 0
    summand1 = int(num//2)
    summand2 = int(num//2) + remainder
    return [summand1, summand2]

output = [0x00] * MESSAGE_LENGTH
output[0] = KEY
output[1] = STATE
output[2] = LRL(user_data[LOWER_RATE_LIMIT])
output[3] = APW(user_data[ATRIAL_PULSE_WIDTH])
output[4] = VPW(user_data[VENTRICULAR_PULSE_WIDTH])
output[5] = fVRP(user_data[VRP])[0]
output[6] = fVRP(user_data[VRP])[1]
output[7] = fARP(user_data[ARP])[0]
output[8] = fARP(user_data[ARP])[1]
output[9] = AA(user_data[ATRIAL_AMPLITUDE]) 
output[10] = VA(user_data[VENTRICULAR_AMPLITUDE])
output[11] = AVD(user_data[AV_DELAY])[0]
output[12] = AVD(user_data[AV_DELAY])[1]



''' SERIAL SETUP '''
def connect(port_name):
    try:
        ser = serial.Serial()
        ser.port = port_name
        ser.baudrate = BAUDRATE
        ser.timeout = TIMEOUT
        return ser
    except:
        return None
    
print(output)
portname = '/dev/cu.usbmodem0000001234561'
ser = connect(portname)

