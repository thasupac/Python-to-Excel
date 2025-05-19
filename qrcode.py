from tkinter import Frame, LabelFrame, Text, ttk, END, messagebox
import PIL.Image
import configparser
import segno

class CreateQRcode(Frame):
    def __init__(self, GUI):

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

        self.qrcodepath = self.config['EXPORTPATH']['qrcode']

        Frame.__init__(self, GUI, width=1500, height=1500)

        def createqr():
            value = (str(text.get('1.0', END))).strip()
            if value:
                qrcode = segno.make_qr(value)
                qrcode.save(f"{self.qrcodepath}/Create_QR_code.png", scale=8)
                qr_path = f"{self.qrcodepath}/Create_QR_code.png"
                qr = PIL.Image.open(qr_path)
                qr.show()
                reset()

            else:
                messagebox.showinfo('Create QR', 'โปรดกรอกข้อมูล')

        def reset():
            text.delete('1.0', END)


        main_frame = LabelFrame(self, text='Create QR code :')
        main_frame.pack()

        text = Text(main_frame, width=20, height=3)
        text.pack(padx=10, pady=10)

        B = ttk.Button(main_frame, text='Create', command=createqr)
        B.pack(ipadx=5, ipady=5)

# from tkinter import *
# gui = Tk()
# gui.title('create QR')
# a = CreateQRcode(gui)
# a.pack()
# gui.mainloop()


