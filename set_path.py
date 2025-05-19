from tkinter import ttk, messagebox, filedialog, Frame, LabelFrame, Label, StringVar
import configparser

#font size
FONT1 = ('Angsana New', 25, 'bold')
FONT2 = ('Angsana New', 18, )
FONT3 = ('Angsana New', 16, )

###---
class set_path_doc(Frame):
    def __init__(self, GUI):
        Frame.__init__(self, GUI, width=1500, height=1500)

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

        #-save path function
        def save_path():
            
            #mc
            path_mc = v_mc_path.get()
            path_mc_machine_sheet = v_mc_sheet.get()

            #spare
            path_spare = v_spare_path.get()
            path_spare_SparePart_sheet = v_Sparepart_sheet.get()
            path_spare_TrackSpare_sheet = v_TrackSpare_sheet.get()
            path_spare_Requestlog_sheet = v_Requestlog.get()

            #tool
            path_tool = v_tool_path.get()
            path_tool_fixture_sheet = v_fixture_sheet.get()
            path_tool_bordProfile_sheet = v_bordProfile_sheet.get()
            path_regTool = v_Reg_tool_sheet.get()
            path_stenCil = v_stencil_sheet.get()

            #mc_down
            path_mcdown = v_machinedowntime_path.get()
            sheet_macdown = v_machinedowntime_sheet.get()



            try:
                #mc
                if path_mc and path_mc_machine_sheet != (''):
                    self.config.set('DATABASE', 'dBmachinePath', path_mc)
                    self.config.set('DATABASE', 'machineSheet', path_mc_machine_sheet)

                #spare
                if path_spare and path_spare_SparePart_sheet and path_spare_TrackSpare_sheet != (''):
                    self.config.set('DATABASE', 'dBsparepath', path_spare)
                    self.config.set('DATABASE', 'sparepartSheet', path_spare_SparePart_sheet)
                    self.config.set('DATABASE', 'trackspareSheet', path_spare_TrackSpare_sheet)
                    self.config.set('DATABASE', 'requestlogSheet', path_spare_Requestlog_sheet)
                
                #tool
                if path_tool and path_tool_fixture_sheet and path_tool_bordProfile_sheet != (''):
                    self.config.set('DATABASE', 'dBtoolPath', path_tool)
                    self.config.set('DATABASE', 'regtoolSheet', path_regTool)
                    self.config.set('DATABASE', 'palletSheet', path_tool_fixture_sheet)
                    self.config.set('DATABASE', 'bordprofileSheet', path_tool_bordProfile_sheet)
                    self.config.set('DATABASE', 'stencilSheet', path_stenCil)

                #mc_down
                if path_mcdown and sheet_macdown != (''):
                    self.config.set('DATABASE', 'dbdowntimepath', path_mcdown)
                    self.config.set('DATABASE', 'downtimesheet', sheet_macdown)

                #save
                with open('config.ini', 'w') as configfile:
                    self.config.write(configfile)
                    messagebox.showinfo('Setting','บันทึก Path สำเร็จแล้ว')

            except Exception as e :
                messagebox.showerror('Setting path',f'Fail to save data : error : {e}')

        #-main frame
        MF = LabelFrame(self)
        MF.pack(padx=10, pady=10)

        ###machine
        MACHINE = LabelFrame(MF, text='dB m/c', font=FONT3)
        MACHINE.grid(row=0, column=0, padx=10, pady=10)

        #-mc
        L = Label(MACHINE, text='dB m/c path :', font=FONT2)
        L.grid(row=0, column=0, padx=10, pady=5)
        v_mc_path = StringVar()
        v_mc_path.set(self.config['DATABASE']['dBmachinePath'])
        E = ttk.Entry(MACHINE, textvariable=v_mc_path, font=FONT2)
        E.grid(row=0, column=1, padx=10, pady=5)

        #mc sheet
        L = Label(MACHINE, text='dB mc sheet :', font=FONT2)
        L.grid(row=1, column=0, padx=20, pady=5)
        v_mc_sheet = StringVar()
        v_mc_sheet.set(self.config['DATABASE']['machineSheet'])
        E = ttk.Entry(MACHINE, textvariable=v_mc_sheet, font=FONT2)
        E.grid(row=1, column=1, padx=20, pady=5)

        #button directory
        def FileDir():
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm")])
            if file_path:
                v_mc_path.set(file_path)

        B = ttk.Button(MACHINE, text='Open', compound='top',command=FileDir)
        B.grid(row=2, column=1, pady=5)

        ###spare
        SPARE = LabelFrame(MF, text='dB spare', font=FONT3)
        SPARE.grid(row=0, column=1, padx=10, pady=10)
        v_spare_path = StringVar()
        v_spare_path.set(self.config['DATABASE']['dBsparepath'])
        L = Label(SPARE, text='dB spare path :', font=FONT2)
        L.grid(row=0, column=0, padx=10, pady=10)
        E = ttk.Entry(SPARE, textvariable=v_spare_path, font=FONT2)
        E.grid(row=0, column=1, padx=10, pady=10)

        #dB SparePart_sheet
        L = Label(SPARE, text='dB SparePart_sheet :', font=FONT2)
        L.grid(row=1, column=0, padx=10, pady=10)
        v_Sparepart_sheet = StringVar()
        v_Sparepart_sheet.set(self.config['DATABASE']['sparepartSheet'])
        E = ttk.Entry(SPARE, textvariable=v_Sparepart_sheet, font=FONT2)
        E.grid(row=1, column=1, padx=10, pady=10)
        
        #dB TrackSpare shee
        L = Label(SPARE, text='dB TrackSpare sheet :', font=FONT2)
        L.grid(row=2, column=0, padx=10, pady=10)
        v_TrackSpare_sheet = StringVar()
        v_TrackSpare_sheet.set(self.config['DATABASE']['trackspareSheet'])
        E = ttk.Entry(SPARE, textvariable=v_TrackSpare_sheet, font=FONT2)
        E.grid(row=2, column=1, padx=10, pady=10)

        #dB Sparelog
        L = Label(SPARE, text='dB Requestlog sheet :', font=FONT2)
        L.grid(row=3, column=0, padx=10, pady=10)
        v_Requestlog = StringVar()
        v_Requestlog.set(self.config['DATABASE']['requestlogSheet'])
        E = ttk.Entry(SPARE, textvariable=v_Requestlog, font=FONT2)
        E.grid(row=3, column=1, padx=10, pady=10)
        
        #button directory
        def FileDir():
            file_path_part = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm")])
            if file_path_part:
                v_spare_path.set(file_path_part)

        B = ttk.Button(SPARE, text='Open', compound='top',command=FileDir)
        B.grid(row=4, column=1, pady=10)

        ###-tool
        TOOL = LabelFrame(MF, text='dB tool', font=FONT3, width=800, height=200)
        TOOL.grid(row=0, column=2, padx=10, pady=10)
        
        #toolpath
        L = Label(TOOL, text='dB tool path :', font=FONT2)
        L.grid(row=0, column=0, padx=10, pady=10)
        v_tool_path = StringVar()
        v_tool_path.set(self.config['DATABASE']['dBtoolPath'])
        E = ttk.Entry(TOOL, textvariable=v_tool_path, font=FONT2)
        E.grid(row=0, column=1, padx=10, pady=10)

        #fixture
        L = Label(TOOL, text='dB fixture sheet :', font=FONT2)
        L.grid(row=1, column=0, padx=10, pady=10)
        v_fixture_sheet = StringVar()
        v_fixture_sheet.set(self.config['DATABASE']['palletSheet'])
        E = ttk.Entry(TOOL, textvariable=v_fixture_sheet, font=FONT2)
        E.grid(row=1, column=1, padx=10, pady=10)

        #bord profile
        L = Label(TOOL, text='dB bord profile sheet :', font=FONT2)
        L.grid(row=2, column=0, padx=10, pady=10)
        v_bordProfile_sheet = StringVar()
        v_bordProfile_sheet.set(self.config['DATABASE']['bordprofileSheet'])
        E = ttk.Entry(TOOL, textvariable=v_bordProfile_sheet, font=FONT2)
        E.grid(row=2, column=1, padx=10, pady=10)

        #regtool
        L = Label(TOOL, text='dB Reg tool sheet :', font=FONT2)
        L.grid(row=3, column=0, padx=10, pady=10)
        v_Reg_tool_sheet = StringVar()
        v_Reg_tool_sheet.set(self.config['DATABASE']['regtoolSheet'])
        E = ttk.Entry(TOOL, textvariable=v_Reg_tool_sheet, font=FONT2)
        E.grid(row=3, column=1, padx=10, pady=10)

        #stencil
        L = Label(TOOL, text='dB stencil sheet :', font=FONT2)
        L.grid(row=4, column=0, padx=10, pady=10)
        v_stencil_sheet = StringVar()
        v_stencil_sheet.set(self.config['DATABASE']['stencilSheet'])
        E = ttk.Entry(TOOL, textvariable=v_stencil_sheet, font=FONT2)
        E.grid(row=4, column=1, padx=10, pady=10)

        #button directory
        def FileDir():
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm")])
            if file_path:
                v_tool_path.set(file_path)

        B = ttk.Button(TOOL, text='Open', compound='top',command=FileDir)
        B.grid(row=5, column=1, pady=10)

        ###machine downtime
        mc_down = LabelFrame(MF, text='db machine downtime', font=FONT3)
        mc_down.grid(row=1, column=0, padx=10, pady=10)

        ##pathmachine downtime
        L = Label(mc_down, text='dB machine downtime :', font=FONT2)
        L.grid(row=0, column=0, padx=10, pady=10)
        v_machinedowntime_path = StringVar()
        v_machinedowntime_path.set(self.config['DATABASE']['dbdowntimepath'])
        E = ttk.Entry(mc_down, textvariable=v_machinedowntime_path, font=FONT2)
        E.grid(row=0, column=1, padx=10, pady=10)

        #machine downtime sheet
        L = Label(mc_down, text='dB machine downtime sheet :', font=FONT2)
        L.grid(row=1, column=0, padx=10, pady=10)
        v_machinedowntime_sheet = StringVar()
        v_machinedowntime_sheet.set(self.config['DATABASE']['downtimesheet'])
        E = ttk.Entry(mc_down, textvariable=v_machinedowntime_sheet, font=FONT2)
        E.grid(row=1, column=1, padx=10, pady=10)

        #button directory
        def FileDir():
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm")])
            if file_path:
                v_machinedowntime_path.set(file_path)

        B = ttk.Button(mc_down, text='Open', compound='top',command=FileDir)
        B.grid(row=2, column=1, pady=10)

        

        #save path
        B = ttk.Button(MF, text='Save', command=save_path)
        B.grid(row=1, column=1, padx=10, pady=10)

class set_path_photo(Frame):
    def __init__(self, GUI):
        Frame.__init__(self, GUI, width=1500, height=1500)

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

        #-save path function
        def save_path_photo():
            #
            path_photo_mc = v_mc_path.get()
            
            #
            path_photo_spare = v_spare_path.get()
            
            #
            path_photo_tool = v_tool_path.get()

            try:
    
                #photo mc
                if path_photo_mc !=(''):
                    self.config.set('DATABASE', 'dbphotomachinepath', path_photo_mc)
                
                #photo spare
                if path_photo_spare !=(''):
                    self.config.set('DATABASE', 'dbphotosparepath', path_photo_spare)
                
                #photo tool
                if path_photo_tool !=(''):
                    self.config.set('DATABASE', 'dbphototoolpath', path_photo_tool)

                #save path photo
                with open('config.ini', 'w') as configfile:
                    self.config.write(configfile)
                    messagebox.showinfo('Setting','บันทึก Path สำเร็จแล้ว')

            except Exception as e :
                messagebox.showerror('Setting path',f'Fail to save data : error : {e}')

        def mcphoto():
            file_path = filedialog.askdirectory()
            if file_path !=(''):
                v_mc_path.set(file_path)

        def sparephoto():
            file_path = filedialog.askdirectory()
            if file_path !=(''):
                v_spare_path.set(file_path)
        def toolphoto():
            file_path = filedialog.askdirectory()
            if file_path !=(''):
                v_tool_path.set(file_path)

        #-main frame
        MF = LabelFrame(self, width=850, height=800)
        MF.pack()

        #-mc
        F = LabelFrame(MF, text='dB m/c', font=FONT3)
        F.grid(row=0, column=0, padx=10)
        v_mc_path = StringVar()
        v_mc_path.set(self.config['DATABASE']['dbphotomachinepath'])
        L = Label(F, text='dB m/c path photo :', font=FONT2)
        L.grid(row=0, column=0, padx=30, pady=10)
        E = ttk.Entry(F, textvariable=v_mc_path, font=FONT2, width=30)
        E.grid(row=0, column=1, padx=10, pady=10)
        B = ttk.Button(F, text='Open', command=mcphoto)
        B.grid(row=0, column=2, padx=10)

        #-spare
        F2 = LabelFrame(MF, text='dB spare', font=FONT3)
        F2.grid(row=1, column=0)
        v_spare_path = StringVar()
        v_spare_path.set(self.config['DATABASE']['dbphotosparepath'])
        L = Label(F2, text='dB spare path photo :', font=FONT2)
        L.grid(row=0, column=0, padx=10, pady=10)
        E = ttk.Entry(F2, textvariable=v_spare_path, font=FONT2, width=30)
        E.grid(row=0, column=1, padx=10)
        B = ttk.Button(F2, text='Open', command=sparephoto)
        B.grid(row=0, column=2, padx=10)
        
        #-tool
        F3 = LabelFrame(MF, text='dB tool', font=FONT3, width=800, height=110)
        F3.grid(row=2, column=0)
        v_tool_path = StringVar()
        v_tool_path.set(self.config['DATABASE']['dbphototoolpath'])
        L = Label(F3, text='dB tool path photo :', font=FONT2)
        L.grid(row=0 ,column=0, padx=10, pady=10)
        E = ttk.Entry(F3, textvariable=v_tool_path, font=FONT2, width=30)
        E.grid(row=0, column=1, padx=10, pady=10)
        B = ttk.Button(F3, text='Open', command=toolphoto)
        B.grid(row=0, column=2, padx=10)

        B = ttk.Button(MF, text='Save', command=save_path_photo)
        B.grid(row=4, column=0, pady=10)

class ConfigmailNotification(Frame):
    def __init__(self, GUI):
        Frame.__init__(self, GUI)

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

        def save_username():
            #
            user_name = v_account.get()
            password = v_password.get()
            receive = v_receive.get()

            try:
                if user_name !=(''):
                    self.config.set('ACCESSOUTLOOK', 'username', user_name)
                if password !=(''):
                    self.config.set('ACCESSOUTLOOK', 'password', password)
                if receive !=(''):
                    self.config.set('ACCESSOUTLOOK', 'sendto', receive)
                
                #save path photo
                with open('config.ini', 'w') as configfile:
                    self.config.write(configfile)
                    messagebox.showinfo('Setting','บันทึกสำเร็จแล้ว')

            except Exception as e :
                messagebox.showerror('Account and notification',f'Fail to save data : error : {e}')

        #-main frame
        MF = LabelFrame(self, width=850, height=800)
        MF.pack(padx=10, pady=10)

        #-account
        F = LabelFrame(MF, text='Account and notification :', font=FONT3)
        F.grid(row=0, column=0, padx=10)
        v_account = StringVar()
        v_account.set(self.config['ACCESSOUTLOOK']['username'])
        L = Label(F, text='Username :', font=FONT2)
        L.grid(row=0, column=0, padx=30, pady=10)
        E = ttk.Entry(F, textvariable=v_account, font=FONT2, width=30)
        E.grid(row=0, column=1, padx=10)

        #password
        v_password = StringVar()
        v_password.set(self.config['ACCESSOUTLOOK']['password'])
        L = Label(F, text='Password :', font=FONT2)
        L.grid(row=1, column=0, padx=30, pady=10)
        E = ttk.Entry(F, textvariable=v_password, font=FONT2, width=30)
        E.grid(row=1, column=1, padx=10)
       

        #receive
        v_receive = StringVar()
        v_receive.set(self.config['ACCESSOUTLOOK']['sendto'])
        L = Label(F, text='Receive :', font=FONT2)
        L.grid(row=2, column=0, padx=30, pady=10)
        E = ttk.Entry(F, textvariable=v_receive, font=FONT2, width=30)
        E.grid(row=2, column=1, padx=10)
        B = ttk.Button(F, text='Save', command=save_username)
        B.grid(row=2, column=2, padx=10)



# from tkinter import *
# gui = Tk()
# setPath = set_path_doc(gui)
# setPath.pack()
# gui.mainloop()



        
        

