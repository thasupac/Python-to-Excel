from tkinter import ttk, Tk, Frame, Menu, BOTH, Toplevel, PhotoImage, Label, LEFT, RIGHT,messagebox

######function center window
def center_windows(w,h):
    ws = GUI.winfo_screenwidth() #screen width
    hs = GUI.winfo_screenheight() #screen height
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    return f'{w}x{h}+{x:.0f}+{y:.0f}'

#หน้าต่างรวมทั้งหมด
GUI = Tk()
GUI.title('Store management v1.0')
win_size = center_windows(800,500)
GUI.geometry(win_size)

#font size
FONT1 = ('Angsana New', 25, 'bold')
FONT2 = ('Angsana New', 18, )
FONT3 = ('Angsana New', 16, )

#create tap
#main tap
Tap = ttk.Notebook(GUI)
Home_Tap = Frame(Tap)
Setting_Tap = Frame(Tap)
Tap.add(Home_Tap, text='Home')
Tap.pack(fill=BOTH, expand=1)


#config tab
s = ttk.Style()
s.configure('TNotebook.Tab',font=(None,9),padding=[25,1])

#main menu
menubar = Menu(GUI)
GUI.config(menu=menubar)

#menu file
menu_file = Menu(menubar, tearoff=0)
menubar.add_cascade(label='File', menu=menu_file)

#menu exportdata
def stencildata():
    from exportdata import ExportDataStencil
    GUI = Toplevel()
    GUI.title('Report stencil data')
    stencil_data = ExportDataStencil(GUI)
    stencil_data.pack()

def downtimedata():
    from exportdata import ExportDatadowntime
    GUI = Toplevel()
    GUI.title('Report downtime data')
    downtime = ExportDatadowntime(GUI)
    downtime.pack()

def toolingdata():
    from exportdata import ExportDatatooling
    GUI = Toplevel()
    GUI.title('Report tooling lists')
    downtime = ExportDatatooling(GUI)
    downtime.pack()

def sparestock():
    from exportdata import ExportDatasparepart
    GUI = Toplevel()
    GUI.title('Spare part stock report')
    downtime = ExportDatasparepart(GUI)
    downtime.pack()

sub_menu_export = Menu(menu_file, tearoff=0)
sub_menu_export.add_command(label='stencil', command=stencildata)
sub_menu_export.add_command(label='fixture')
sub_menu_export.add_command(label='bordprofile')
sub_menu_export.add_command(label='machine downtime', command=downtimedata)
sub_menu_export.add_command(label='Tooling lists', command=toolingdata)
sub_menu_export.add_command(label='Spare stock', command=sparestock)
menu_file.add_cascade(label='Export data', menu=sub_menu_export)

##config database
def doc_path():
    from set_path import set_path_doc
    GUI = Toplevel()
    GUI.geometry('600x600')
    GUI.title('Setting path')
    doc_path = set_path_doc(GUI)
    doc_path.pack()
    
##config photo database
def photo_path():
    from set_path import set_path_photo
    GUI = Toplevel()
    GUI.geometry('600x600')
    GUI.title('Setting path')
    doc_path = set_path_photo(GUI)
    doc_path.pack()

def account_notification():
    from set_path import ConfigmailNotification
    GUI = Toplevel()
    GUI.title('Setting path')
    acount_noti = ConfigmailNotification(GUI)
    acount_noti.pack()

sub_menu = Menu(menu_file, tearoff=0)
sub_menu.add_command(label='config database', command=doc_path)
sub_menu.add_command(label='config photo database', command=photo_path)
sub_menu.add_command(label='config account and notification', command=account_notification)
menu_file.add_cascade(label='Config', menu=sub_menu)

#menu exit
menu_file.add_separator()
menu_file.add_command(label='Exit', accelerator='Ctr + Q', command=lambda: GUI.quit())
GUI.bind('<Control-q>', lambda x:GUI.quit())

#ทำเมนูเป็น icon
#machine
def WindowMachine():
    from mcpage import MCreg, MCview
    GUI.withdraw()
    mcgui = Toplevel()
    mcgui.geometry(win_size)
    mcgui.title('Machine')

    #แยก Tap
    mc = ttk.Notebook(mcgui)
    mcreg = Frame(mc)
    mcviews = Frame(mc)
    mc.add(mcreg, text='Register')
    mc.add(mcviews, text='Views')
    mc.pack(fill=BOTH, expand=1)

    #insert to tap
    mcregTap = MCreg(mcreg)
    mcregTap.pack()

    mcviewsTap = MCview(mcviews)
    mcviewsTap.pack()  

    #main menu
    menubar = Menu(mcgui)
    mcgui.config(menu=menubar)

    def BackHomE():
        MCreg.StartmCSwitchPagE(GUI) #GUI
    menubar.add_cascade(label='Home', command=BackHomE)

    # แก้ปัญหาปิดครั้งแรก
    def wd():
        check = messagebox.askyesno('Store management', 'Do you want to close?')
        if check:
            GUI.destroy()
            GUI.quit()
    mcgui.protocol("WM_DELETE_WINDOW", wd)

MC = Frame(Home_Tap)
MC.place(x=100, y=50)
icon_mc = PhotoImage(file='machine.png')
BMC = ttk.Button(MC, text='Machine', image=icon_mc, compound='top', command=WindowMachine)
BMC.pack(ipadx=50, ipady=20)
    
#spare part
def WindowSpare():
    from partpage import PartReg, PartViews
    from mcpage import MCreg
    GUI.withdraw()
    spgui = Toplevel()
    spgui.geometry(win_size)
    spgui.title('Spare part')

    #แยก tap
    sp = ttk.Notebook(spgui)
    spreg = Frame(sp)
    spviews = Frame(sp)
    sp.add(spreg, text='Register')
    sp.add(spviews, text='Views')
    sp.pack(fill=BOTH, expand=1)

    #insert to tap
    spregTap = PartReg(spreg)
    spregTap.pack()
    spviewsTap = PartViews(spviews)
    spviewsTap.pack()

    #main menu
    menubar = Menu(spgui)
    spgui.config(menu=menubar)

    def BackHomE():
        MCreg.StartmCSwitchPagE(GUI)
    menubar.add_cascade(label='Home', command=BackHomE)

    #แก้ปัญหาปิดครั้งแรก
    def wd():
        check = messagebox.askyesno('Store management', 'Do you want to close?')
        if check:
            GUI.destroy()
            GUI.quit()
    spgui.protocol("WM_DELETE_WINDOW", wd)

SP = Frame(Home_Tap)
SP.place(x=300,y=50)
icon_sp = PhotoImage(file='spare.png')
BSP = ttk.Button(SP, text='Spare', image=icon_sp, compound='top', command=WindowSpare)
BSP.pack(ipadx=50, ipady=20)

#tool
def WindowTool():
    from mcpage import MCreg
    GUI.withdraw()
    toolgui = Toplevel()
    toolgui.geometry(win_size)
    toolgui.title('Tooling')

    #แยก tap
    tool = ttk.Notebook(toolgui)
    maintool = Frame(tool)
    tool.add(maintool)
    tool.pack(fill=BOTH, expand=1)

    #maintool
    def regmanagetool():
        from toolpage import RegisterTool, ViewsTool
        GUI.withdraw()
        managetool = Toplevel()
        managetool.geometry('1000x500')
        managetool.title('Manage tooling')

        managetooltap = ttk.Notebook(managetool)
        regtool = Frame(managetooltap)
        viewstool = Frame(managetooltap)
        managetooltap.add(regtool, text='Register')
        managetooltap.add(viewstool, text='Views')
        managetooltap.pack(fill=BOTH, expand=1)

        regtool = RegisterTool(regtool)
        regtool.pack()
        viewstool = ViewsTool(viewstool)
        viewstool.pack()

    managetool = Frame(maintool)
    managetool.place(x=100, y=50)
    Bmanagetool = ttk.Button(managetool, text='Manage', image=icon_managetool, compound='top', command=regmanagetool)
    Bmanagetool.pack(ipadx=50, ipady=20)

    #fixture
    def fixturetool():
        from toolpage import fixWithdraw, fixPageReturn
        GUI.withdraw()
        fixturegui = Toplevel()
        fixturegui.geometry('1000x500')
        fixturegui.title('Fixture tooling')

        fixturetap = ttk.Notebook(fixturegui)
        withdraw = Frame(fixturetap)
        receive = Frame(fixturetap)
        fixturetap.add(withdraw, text='Withdraw')
        fixturetap.add(receive, text='Receive')
        fixturetap.pack(fill=BOTH, expand=1)

        with_fixture = fixWithdraw(withdraw)
        with_fixture.pack()
        return_fixture = fixPageReturn(receive)
        return_fixture.pack()

    fixture = Frame(maintool)
    fixture.place(x=300,y=50)
    Bfixturetool = ttk.Button(fixture, text='Fixture', image=icon_fixture, compound='top', command=fixturetool)
    Bfixturetool.pack(ipadx=50, ipady=20)

    #bordprofile
    def bordprofiletool():
        from toolpage import bordwith, bordReturn
        GUI.withdraw()
        bordprofilegui = Toplevel()
        bordprofilegui.geometry('1000x500')
        bordprofilegui.title('Bordprofile tooling')

        bordprofiletap = ttk.Notebook(bordprofilegui)
        withdraw = Frame(bordprofiletap)
        receive = Frame(bordprofiletap)
        bordprofiletap.add(withdraw, text='Withdraw')
        bordprofiletap.add(receive, text='Receive')
        bordprofiletap.pack(fill=BOTH, expand=1)

        with_bord = bordwith(withdraw)
        with_bord.pack()
        return_bord = bordReturn(receive)
        return_bord.pack()

    bordprofile = Frame(maintool)
    bordprofile.place(x=500, y=50)
    Bbordprofile = ttk.Button(bordprofile, text='Bordprofile', image=icon_bord, compound='top', command=bordprofiletool)
    Bbordprofile.pack(ipadx=50, ipady=20)

    #stencil
    def stenciltool():
        from toolpage import stenCilwith, stenCilre
        GUI.withdraw()
        stenciltoolgui = Toplevel()
        stenciltoolgui.geometry('1000x500')
        stenciltoolgui.title('Stencil tooling')

        stenciltooltap = ttk.Notebook(stenciltoolgui)
        withdraw = Frame(stenciltooltap)
        receive = Frame(stenciltooltap)
        stenciltooltap.add(withdraw, text='Withdraw')
        stenciltooltap.add(receive, text='Receive')
        stenciltooltap.pack(fill=BOTH, expand=1)

        sten_cil_with_tab = stenCilwith(withdraw)
        sten_cil_with_tab.pack()
        sten_cil_re_tap = stenCilre(receive)
        sten_cil_re_tap.pack()


    stencil = Frame(maintool)
    stencil.place(x=100, y=200)
    Bstencil = ttk.Button(stencil, text='Stencil', image=icon_stencil, compound='top', command=stenciltool)
    Bstencil.pack(ipadx=50, ipady=20)

    #main menu
    menubar = Menu(toolgui)
    toolgui.config(menu=menubar)

    def BackHomE():
        MCreg.StartmCSwitchPagE(GUI)
    menubar.add_cascade(label='Home', command=BackHomE)

    #แก้ปัญหาปิดครั้งแรก
    def wd():
        check = messagebox.askyesno('Store management', 'Do you want to close?')
        if check:
            GUI.destroy()
            GUI.quit()
    toolgui.protocol("WM_DELETE_WINDOW", wd)

Tool = Frame(Home_Tap)
Tool.place(x=500, y=50)
icon_tool = PhotoImage(file='tool.png')
BTool = ttk.Button(Tool, text='Tool', image=icon_tool, compound='top',command=WindowTool)
BTool.pack(ipadx=50, ipady=20)

#icon
icon_managetool = PhotoImage(file='managetool.png')
icon_fixture = PhotoImage(file='fixture.png')
icon_bord = PhotoImage(file='bord.png')
icon_stencil = PhotoImage(file='stencil.png')

#request
def WindowRequest():
    from requestpage import Request, Requestview
    GUI.withdraw()
    rqgui = Toplevel()
    rqgui.geometry(win_size)
    rqgui.title('Request Spare part')

    #แยก tap
    rq = ttk.Notebook(rqgui)
    rqspare = Frame(rq)
    rqspareview = Frame(rq)
    rq.add(rqspare, text='Register')
    rq.add(rqspareview, text='Spare views')
    rq.pack(fill=BOTH, expand=1)

    #insert to tap
    rqTap = Request(rqspare)
    rqTap.pack()
    rqvTap = Requestview(rqspareview)
    rqvTap.pack()

    #main menu
    menubar = Menu(rqgui)
    rqgui.config(menu=menubar)

    def BackHomE():
        from mcpage import MCreg
        MCreg.StartmCSwitchPagE(GUI)
    menubar.add_cascade(label='Home', command=BackHomE)

    #แก้ปัญหาปิดครั้งแรก
    def wd():
        check = messagebox.askyesno('Store management', 'Do you want to close?')
        if check:
            GUI.destroy()
            GUI.quit()
    rqgui.protocol("WM_DELETE_WINDOW", wd)

rq = Frame(Home_Tap)
rq.place(x=100, y=200)
icon_rq = PhotoImage(file='request.png')
RQ = ttk.Button(rq, text='Request', image=icon_rq, compound='top', command=WindowRequest)
RQ.pack(ipadx=50, ipady=20)

#downtime
def Windowndowntime():
    from mc_down import InformMachinedown, ActionMachinedown
    GUI.withdraw()
    machine_downtime = Toplevel()
    machine_downtime.geometry('1000x500')
    machine_downtime.title('Machine downtime')

    machine_downtime_tap = ttk.Notebook(machine_downtime)
    inform = Frame(machine_downtime_tap)
    action = Frame(machine_downtime_tap)
    machine_downtime_tap.add(inform, text='Inform')
    machine_downtime_tap.add(action, text='Action')
    machine_downtime_tap.pack(fill=BOTH, expand=1)

    mc_down_inform = InformMachinedown(inform)
    mc_down_inform.pack()
    mc_down_action = ActionMachinedown(action)
    mc_down_action.pack()

    #main menu
    menubar = Menu(machine_downtime)
    machine_downtime.config(menu=menubar)

    def BackHomE():
        from mcpage import MCreg
        MCreg.StartmCSwitchPagE(GUI)
    menubar.add_cascade(label='Home', command=BackHomE)

    #แก้ปัญหาปิดครั้งแรก
    def wd():
        check = messagebox.askyesno('Store management', 'Do you want to close?')
        if check:
            GUI.destroy()
            GUI.quit()
    machine_downtime.protocol("WM_DELETE_WINDOW", wd)
   
MCdown = Frame(Home_Tap)
MCdown.place(x=300, y=200)
icon_dmc = PhotoImage(file='down.png')
MCD = ttk.Button(MCdown, text='Downtime', image=icon_dmc, compound='top', command=Windowndowntime)
MCD.pack(ipadx=50, ipady=20)

########################
########################
########################
######qr code generate
def qr_gen():
    from qrcode import CreateQRcode
    GUI = Toplevel()
    GUI.title('Setting path')
    qr_create = CreateQRcode(GUI)
    qr_create.pack()
    # messagebox.showinfo('waited', 'รอสักครู่ครับ')

qrgen = Frame(Home_Tap)
qrgen.place(x=500, y=200)
icon_qr = PhotoImage(file='qrcode.png')
qrg = ttk.Button(qrgen, text='QR code', image=icon_qr, compound='top', command=qr_gen)
qrg.pack(ipadx=50, ipady=20)

########################
########################
########################
######qr code generate

#report
def Sendreport():
    from reportproblem import ReportProblem
    GUI.withdraw()
    Reportgui = Toplevel()
    Reportgui.geometry(win_size)
    Reportgui.title('Report to developer')

    rpTap = ReportProblem(Reportgui)
    rpTap.pack()

    #main menu
    menubar = Menu(Reportgui)
    Reportgui.config(menu=menubar)

    def BackHomE():
        from mcpage import MCreg
        MCreg.StartmCSwitchPagE(GUI)
    menubar.add_cascade(label='Home', command=BackHomE)

    #แก้ปัญหาปิดครั้งแรก
    def wd():
        check = messagebox.askyesno('Store management', 'Do you want to close?')
        if check:
            GUI.destroy()
            GUI.quit()
    Reportgui.protocol("WM_DELETE_WINDOW", wd)
   
REB = ttk.Button(GUI, text='Report', command=Sendreport)
REB.pack(side=RIGHT, padx=1, pady=1)

#devolop by
L = Label(GUI, text='Developed by Thanongsak SU')
L.pack(side=LEFT)

GUI.mainloop()