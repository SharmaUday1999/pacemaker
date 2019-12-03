# example data
user_data = [150, 1.2, 1.3, 1.2, 1.4, 187, 200, 300]

# change values to INDEX of param in CSV
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

# lower rate lim (conv to hex)
def LRL(LOWER_RATE_LIMIT):
    return hex(LOWER_RATE_LIMIT)

# atrial pulse width (conv to hex)
def APW(num):
    return hex(int(num*100))

# ventricular pulse width (conv to hex)
def VPW(num):
    return hex(int(num*100))
    
# VRP (split into two summands)
def fVRP(num):
    remainder = 1 if (num % 2 == 1) else 0
    summand1 = hex(int(num//2))
    summand2 = hex(int(num//2) + remainder)
    return [summand1, summand2]

# ARP
def fARP(num):
    remainder = 1 if (num % 2 == 1) else 0
    summand1 = hex(int(num//2))
    summand2 = hex(int(num//2) + remainder)
    return [summand1, summand2]

# atrial amplitude
def AA(num):
    return hex(int(num*10))

# ventricular amplitude
def VA(num):
    return hex(int(num*10))

# AV delay
def AVD(num):
    remainder = 1 if (num % 2 == 1) else 0
    summand1 = hex(int(num//2))
    summand2 = hex(int(num//2) + remainder)
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

print(output)


