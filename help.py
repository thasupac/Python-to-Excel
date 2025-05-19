from tkinter import Frame, Tk, LabelFrame, Label, PhotoImage
from PIL import Image, ImageTk
import tkinter as tk

#font size
FONT1 = ('Angsana New', 25, 'bold')
FONT2 = ('Angsana New', 18, )
FONT3 = ('Angsana New', 16, )

class Help(Frame):
    def __init__ (self, GUI):
        Frame.__init__(self, GUI, width=1500, height=1500)

        #Location of measurment
        meas = LabelFrame(self, text='Step to measurement', width=100, height=100)
        meas.pack(padx=10, pady=10, side='left')

        #step1
        L = Label(meas, text='Step 1', font=FONT2)
        L.grid(row=0, column=0)

        #photo step 1
        PL = Label(meas)
        PL.grid(row=1, column=0, padx=10, pady=10)
        image = Image.open('1_setzero.png')
        image = image.resize((200,140))
        photo = ImageTk.PhotoImage(image)
        PL.config(image=photo)
        PL.image = photo

        #step2
        L = tk.Label(meas, text='Step 2', font=FONT2)
        L.grid(row=0, column=1)

        #photo step2
        PL2 = Label(meas)
        PL2.grid(row=1, column=1, padx=10, pady=10)
        image = Image.open('2_point.png')
        image = image.resize((200,140))
        photo = ImageTk.PhotoImage(image)
        PL2.config(image=photo)
        PL2.image = photo

        #criteria
        cri = LabelFrame(self, text='Example of criteria', width=100, height=100)
        cri.pack(padx=10, pady=10, side='right')

        #dent
        dent = Label(cri, text='Dent', font=FONT2)
        dent.grid(row=0, column=0, padx=10, pady=10)
        
        #photo dent
        p_dent = Label(cri)
        p_dent.grid(row=1, column=0, padx=10, pady=10)
        image = Image.open('3_dent.png')
        image = image.resize((200,140))
        photo = ImageTk.PhotoImage(image)
        p_dent.config(image=photo)
        p_dent.image = photo

        #scratched
        scratched = Label(cri, text='Scratched', font=FONT2)
        scratched.grid(row=0, column=1, padx=10, pady=10)

        #photo scratched
        p_scratched = Label(cri)
        p_scratched.grid(row=1, column=1, padx=10, pady=10)
        image = Image.open('4_scratched.png')
        image = image.resize((200,140))
        photo = ImageTk.PhotoImage(image)
        p_scratched.config(image=photo)
        p_scratched.image = photo

        #aperture
        L = Label(cri, text='Aperture', font=FONT2)
        L.grid(row=2, column=0)

        #photo scratched
        P_aperture = Label(cri)
        P_aperture.grid(row=3, column=0, padx=10, pady=10)
        image = Image.open('5_aperture.png')
        image = image.resize((200,140))
        photo = ImageTk.PhotoImage(image)
        P_aperture.config(image=photo)
        P_aperture.image = photo

# GUI = Tk()
# CallHelp = Help(GUI)
# CallHelp.pack()
# GUI.mainloop()