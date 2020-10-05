from datetime import date
today = date.today()
print("Today's date:", today)

import time
t = time.localtime()
current_time = time.strftime("%H:%M:%S", t)
print(current_time)

from datetime import datetime
now = datetime.now()
print (now.strftime("%Y-%m-%d %H:%M:%S"))