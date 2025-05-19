from tkinter import ttk, messagebox, Frame, END, LabelFrame, Label, Text, StringVar
from send_mail import SendMail
Report = SendMail()

#font
###---Font
FONT1 = ('Angsana New',25,'bold')
FONT2 = ('Angsana New',18)
FONT3 = ('Angsana New',12)

#techname
techname = ['Thanongsak Su',
            'Khunakorn R',
            'Pichet T',
            'Wasan R',
            'Don P',
            'Somchai L',
            'Adirek C',
            'Sangworn D',
            'Pratchaya S',
            'Supot P',
            'Kriangsak H',
            'Anan C',
            'Thanatorn K',
            'Anong J',
            'Thanongsak D',
            'Apicha K',
            'Sompong L']

class ReportProblem(Frame):
    def __init__(self, GUI):
        Frame.__init__(self, GUI, width=1500, height=1500)

        #sent report
        def Sendreport():
            desc = Entryreport.get('1.0', END)
            line = v_line.get()
            reportby = v_reportby.get()
            if desc and line and reportby:
                Report.StartSendReport(desc,line,reportby)
                reset()
            else:
                messagebox.showinfo('Report problem', 'Please fill data!')
        def reset():
            Entryreport.delete('1.0', END)
            v_line.set('')
            v_reportby.set('')
            

        # #mainframe
        F = LabelFrame(self, text='Report form', font=FONT3, width=800, height=500)
        F.pack(padx=5, pady=5)

        #desc
        L = Label(F, text='Descriptions:', font=FONT2)
        L.grid(row=0, column=0, pady=5, sticky='e')

        #ENTRY
        Entryreport = Text(F, width=40, height=10)
        Entryreport.grid(row=0, column=1, padx=10, pady=10)

        #line
        L = Label(F, text='     Line :', font=FONT2)
        L.grid(row=1, column=0, pady=5, sticky='e')
        v_line = StringVar()
        E = ttk.Combobox(F, textvariable=v_line, font=FONT2, values=['BLD4#2','BLD5#10','BLD6#15'], state='readonly')
        E.grid(row=1, column=1, pady=5)

        #report by
        L = Label(F, text='Report by :', font=FONT2)
        L.grid(row=2, column=0, pady=5, sticky='e')
        v_reportby = StringVar()
        E = ttk.Combobox(F, textvariable=v_reportby, font=FONT2, values= techname, state='readonly')
        E.grid(row=2, column=1, pady=5)

        #button
        B = ttk.Button(self, text='Report', command=Sendreport)
        B.pack(ipadx=10, ipady=10)


# gui = Tk()
# gui.title('Report')
# gui.geometry('500x500')
# a = ReportProblem(gui)
# a.pack()
# gui.mainloop()
