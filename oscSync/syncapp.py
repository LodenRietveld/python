import time
import argparse
import math

import OSC

send_address = 'localhost', 8000

c = OSC.OSCClient()
c.connect(send_address)
time_start = time.clock()

try:
    while True:
        if((time_start + 2) - time.clock() < 0.00001):
            time_start = time.clock()
            msg = OSC.OSCMessage()
            msg.setAddress("/keyframe")
            c.send(msg)

except KeyboardInterrupt:
    pass
