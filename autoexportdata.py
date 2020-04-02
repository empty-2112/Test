import time
import datetime
import schedule
import adddatatoexcel
import shutil

original = r'C:\Users\Admin\Desktop\Howtocode\Template.xlsx'
target = r'C:\Users\Admin\Desktop\Howtocode\Book1.xlsx'

shutil.copyfile(original, target)

#schedule.every().day.at("00:00").do(adddatatoexcel.exportdata)
start = datetime.datetime.strptime("2020-03-20", "%Y-%m-%d")
end = datetime.datetime.strptime("2020-03-30", "%Y-%m-%d")
date_array = \
    (start + datetime.timedelta(days=x) for x in range(0, (end-start).days))
for i in date_array:
    print(i)
    adddatatoexcel.exportdata(i)
    
#for i in data:
   #adddatatoexcel.exportdata(i)
#while True: 
  
    # Checks whether a scheduled task  
    # is pending to run or not         
    #schedule.run_pending() 
    #time.sleep(1)
