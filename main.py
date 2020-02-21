from tkinter import *
from datetime import datetime
import time
import win32com.client as wincl

AMTimes = ["00","01","02","03","04","05","06","07","08","09","10","11"]
times = {
  "00": 0,
  "01": 1,
  "02": 2,
  "03": 3,
  "04": 4,
  "05": 5,
  "06": 6,
  "07": 7,
  "08": 8,
  "09": 9,
  "10": 10,
  "11": 11,
  "12": 12,
  "13": 13,
  "14": 14,
  "15": 15,
  "16": 16,
  "17": 17,
  "18": 18,
  "19": 19,
  "20": 20,
  "21": 21,
  "22": 22,
  "23": 23,
}


########## Functions ##############
def start_clicked():
        lbl3.grid(column=0, row=3)
        time_to_end = "Time left: "
        Hour = str(txt1.get())
        Minute = str(txt2.get())
        now = datetime.now()
        cur_hr = str(now).split()[1].split(":")[0]
        cur_min = str(now).split()[1].split(":")[1]
        if cur_hr in AMTimes:
            addn = "AM"
            num = cur_hr
        else:
            if cur_hr == "12":
                num = "12"
            else:
                num = str(times[cur_hr] - 12)
            addn = "PM" 
        if times[Hour] < times[cur_hr]:
            hr_to_deadline = str(times[Hour] + (24 - times[cur_hr]))
        else:
            hr_to_deadline = str(times[Hour]-times[cur_hr])
        cur_min_val = int(cur_min)
        Minute_val = int(Minute)
        if Minute_val >= cur_min_val:
            min_to_deadline = str(Minute_val - cur_min_val)
        else:
            val1 = 60 - cur_min_val
            val2 = Minute_val
            main_val = val1 + val2
            hr_to_deadline = str(int(hr_to_deadline) - 1)
            min_to_deadline = str(main_val)
        time_to_end += hr_to_deadline + " hrs. and " + min_to_deadline + " mins."
        lbl4 = Label(window, text=time_to_end)
        lbl4.grid(column=0, row=4)
        if int(cur_min) % 10 == 0:
            time_str_1 = "The time is " + num + ":" + cur_min + " " + addn + "."
            time_str_2 = "You have " + hr_to_deadline + " hours and " + min_to_deadline + " minutes to set deadline." 
            speak = wincl.Dispatch("SAPI.SpVoice")
            speak.Speak(time_str_1)
            speak.Speak(time_str_2)
        if hr_to_deadline == "0" and min_to_deadline == "0":
            speak = wincl.Dispatch("SAPI.SpVoice")
            speak.Speak("The deadline is reached. You may continue doing other work.")
            sys.exit()
        window.after(50000, start_clicked)
            

def stop_clicked():
    sys.exit()

######### Main ####################

window = Tk()

window.title("Time Callout")
window.geometry('320x130')
lbl1 = Label(window, text="Enter deadline hour (in HH format)")
lbl1.grid(column=0, row=0)
lbl2 = Label(window, text="Enter deadline minute (MM)")
lbl2.grid(column=0, row=1)
lbl3 = Label(window, text="Callout is active!")
txt1 = Entry(window,width=10)
txt1.grid(column=1, row=0)
txt2 = Entry(window,width=10)
txt2.grid(column=1, row=1)
start_btn = Button(window, text="Start", command=start_clicked)
start_btn.grid(column=0, row=2)
stop_btn = Button(window, text="Stop", command=stop_clicked)
stop_btn.grid(column=1, row=2)
window.mainloop()
