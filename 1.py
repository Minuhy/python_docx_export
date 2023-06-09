from tkinter import *
import time


def onGo():
    t.insert('end', '开始\n')
    for i in range(50):
        t.insert(1.0, 'a_' + str(i) + '\n')
        time.sleep(0.1)
        t.update()


root = Tk()
t = Text(root)
t.pack()
goBtn = Button(text="Go!", command=onGo)
goBtn.pack()
root.mainloop()
