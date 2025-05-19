from accessOutlookEmail import create_account
from accessOutlookEmail import send_email
from tkinter import messagebox,Toplevel,Label
import threading
import configparser

FONT = ('Angsana New',18)

class SendMail:
    def __init__(self):
        #config
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

        #Access outlook
        self.username = self.config['ACCESSOUTLOOK']['username']
        self.password = self.config['ACCESSOUTLOOK']['password']

        self.sendto = self.config['ACCESSOUTLOOK']['sendto'].split(',')
        self.dow_act = self.config['ACCESSOUTLOOK']['dow_act'].split(',')
        
    ##--progress
    def Toolsongmail(self, model, modelNum, customer, clsaaTypes, qtyr, unit, line, regBy, comment, check_photo_save):
        def center_windows(w,h):
            ws = window.winfo_screenwidth() #screen width
            hs = window.winfo_screenheight() #screen height
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            return f'{w}x{h}+{x:.0f}+{y:.0f}'
        
        window = Toplevel()
        win_size = center_windows(200,100)
        window.geometry(win_size)
        window.title('Progress functions')
        label = Label(window, text = 'กำลังส่งข้อมูล...', font=FONT)
        label.pack()
        window.update()
        thread = threading.Thread(target = SendMail.ToolNotify(self, model, modelNum, customer, clsaaTypes, qtyr, unit, line, regBy, comment, check_photo_save))
        thread.start()
        window.destroy()
        window.update()

    def ToolNotify(self, model, modelNum, customer, clsaaTypes, qtyr, unit, line, regBy, comment, check_photo_save):
        subject = 'Tooling management notification'
        body = (f'''Dear all,
                Please add identify of tooling as below.
                Model : {model}
                Model number : {modelNum}
                Customer : {customer}
                Type : {clsaaTypes}
                Quantity : {qtyr} {unit}
                Line : {line}
                Reg by : {regBy}
                Comment : {comment}''')
        
        to = [self.sendto]
        for noti in to:
            try:
                account = create_account(self.username, self.password)
                attachments = []
                if check_photo_save != (''):
                    with open(f'{check_photo_save}', 'rb') as f:
                        content = f.read()
                    attachments.append((f'{modelNum}.{check_photo_save.split('.')[-1]}', content))
                send_email(account, subject, body, noti, attachments)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email via outlook [error : {e}]")

     ##--progress
    def Machinesongmail(self, eqname, eqdesc, brand, serial, ca, ora, wi, form, line, service, pms, dcc, comment,photo_path_check):
        def center_windows(w,h):
            ws = window.winfo_screenwidth() #screen width
            hs = window.winfo_screenheight() #screen height
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            return f'{w}x{h}+{x:.0f}+{y:.0f}'
        
        window = Toplevel()
        win_size = center_windows(200,100)
        window.geometry(win_size)
        window.title('Progress functions')
        label = Label(window, text = 'กำลังส่งข้อมูล...', font=FONT)
        label.pack()
        window.update()
        thread = threading.Thread(target = SendMail.MachineNotify(self, eqname, eqdesc, brand, serial, ca, ora, wi, form, line, service, pms, dcc, comment,photo_path_check))
        thread.start()
        window.destroy()
        window.update()


    def MachineNotify(self, eqname, eqdesc, brand, serial, ca, ora, wi, form, line, service, pms, dcc, comment,photo_path_check):
        subject = 'Machine management notification'
        body = (f'''Dear all,
                Please check new equipment which register as below.
                Equipment name : {eqname}
                Equipment descriptions : {eqdesc}
                Brand : {brand}
                Serial number : {serial}
                CA no : {ca}
                ORA# : {ora}
                WI : {wi}
                Form : {form}
                Line : {line}
                Service : {service}
                PMS : {pms}
                DCC : {dcc}
                comment : {comment}''')

        to = [self.sendto]
        for noti in to:
            try:
                account = create_account(self.username, self.password)
                attachments = []
                if photo_path_check != (''):
                    with open(f'{photo_path_check}', 'rb') as f:
                        content = f.read()
                    attachments.append((f'{eqname}.{photo_path_check.split('.')[-1]}', content))
                send_email(account, subject, body, noti, attachments)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email via outlook [error : {e}]")

    def Sparesongmail(self,serialnum, partname, machine, line, quantity, unit, requestby, comment, attachfile):
        def center_windows(w,h):
            ws = window.winfo_screenwidth() #screen width
            hs = window.winfo_screenheight() #screen height
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            return f'{w}x{h}+{x:.0f}+{y:.0f}'
        
        window = Toplevel()
        win_size = center_windows(200,100)
        window.geometry(win_size)
        window.title('Progress functions')
        label = Label(window, text = 'กำลังส่งข้อมูล...', font=FONT)
        label.pack()
        window.update()
        thread = threading.Thread(target = SendMail.SpareNotify(self, serialnum, partname, machine, line, quantity, unit, requestby, comment, attachfile))
        thread.start()
        window.destroy()
        window.update()
 
    
    def SpareNotify(self,serialnum, partname, machine, line, quantity, unit, requestby, comment, attachfile):
        subject = 'Spare part request from ME team'
        
        body = (f'''Dear all,
                Please review the pare detail as below.
                Serial number : {serialnum}
                Part name : {partname}
                Machine : {machine}
                Building : {line}
                Quantity : {quantity} {unit}
                Requester : {requestby}
                Comment : {comment}''')
        to = [self.sendto]
        for noti in to:
            try:
                account = create_account(self.username, self.password)
                attachments = []
                if attachfile != (''):
                    with open(f'{attachfile}', 'rb') as f:
                        content = f.read()
                    attachments.append((f'{partname}.{attachfile.split('.')[-1]}', content))
                send_email(account, subject, body, noti, attachments)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email via outlook [error : {e}]")

    def Stencilalert(self,modelnum,RebY,desc):
        def center_windows(w,h):
            ws = window.winfo_screenwidth() #screen width
            hs = window.winfo_screenheight() #screen height
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            return f'{w}x{h}+{x:.0f}+{y:.0f}'
        
        window = Toplevel()
        win_size = center_windows(200,100)
        window.geometry(win_size)
        window.title('Progress functions')
        label = Label(window, text = 'กำลังส่งข้อมูล...', font=FONT)
        label.pack()
        window.update()
        thread = threading.Thread(target = SendMail.StencilalertNotify(self,modelnum,RebY,desc))
        thread.start()
        window.destroy()
        window.update()


    def StencilalertNotify(self,modelnum,RebY,desc):
        subject = 'Tooling alert'
        body = (f'''Dear all,
                Please review the the tooling detail as below.
                Stencil number : {modelnum}
                Sender by : {RebY}
                Comment : {desc}''')
        to = [self.sendto]
        for noti in to:
            try:
                account = create_account(self.username, self.password)
                attachments = []
                send_email(account, subject, body, noti, attachments)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email via outlook [error : {e}]")

    def StartSendReport(self, desc, line, reportby):
        def center_windows(w,h):
            ws = window.winfo_screenwidth() #screen width
            hs = window.winfo_screenheight() #screen height
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            return f'{w}x{h}+{x:.0f}+{y:.0f}'
        
        window = Toplevel()
        win_size = center_windows(200,100)
        window.geometry(win_size)
        window.title('Progress functions')
        label = Label(window, text = 'กำลังส่งข้อมูล...', font=FONT)
        label.pack()
        window.update()
        thread = threading.Thread(target = SendMail.Sendreport(self, desc, line, reportby))
        thread.start()
        window.destroy()
        window.update()
        messagebox.showinfo('Report problem', 'Report problem, complete..')
    def Sendreport(self, desc, line, reportby):
        subject = 'Report problem'
        body = (f'''Dear team,
                Please review problem as below.
                Descriptions : {desc}
                Line : {line}
                Sender by : {reportby}''')
        to = [self.sendto]
        for noti in to:
            try:
                account = create_account(self.username, self.password)
                attachments = []
                send_email(account, subject, body, noti, attachments)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email via outlook [error : {e}]")
    
    ###machine downtime
    def Startmachinedowntime(self, timestamp, machine, inform, problems, location):
        def center_windows(w,h):
            ws = window.winfo_screenwidth() #screen width
            hs = window.winfo_screenheight() #screen height
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            return f'{w}x{h}+{x:.0f}+{y:.0f}'
        
        window = Toplevel()
        win_size = center_windows(200,100)
        window.geometry(win_size)
        window.title('Progress functions')
        label = Label(window, text = 'กำลังส่งข้อมูล...', font=FONT)
        label.pack()
        window.update()
        thread = threading.Thread(target = SendMail.Machinedowntime(self, timestamp, machine, inform, problems, location))
        thread.start()
        window.destroy()
        window.update()

    def Machinedowntime(self, timestamp, machine, inform, problems, location):
        subject = 'Machine downtime notification'
        body = (f'''Dear all,
                Please check the machine downtime.
                Time occur : {timestamp}
                Machine : {machine}
                Inform by : {inform}
                Problem : {problems}
                Line : {location}
                Status : Open downtime ''')
       
        to = [self.sendto]
        # to = [self.dow_act]
        for noti in to:
            try:
                account = create_account(self.username, self.password)
                attachments = []
                send_email(account, subject, body, noti, attachments)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email via outlook [error : {e}]")
        

    ###############finished downtime
    def Startfinishedmachinedowntime(self, timestamp, machine, inform, problems, action_fix, action_by, notes,location):
        def center_windows(w,h):
            ws = window.winfo_screenwidth() #screen width
            hs = window.winfo_screenheight() #screen height
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            return f'{w}x{h}+{x:.0f}+{y:.0f}'
        
        window = Toplevel()
        win_size = center_windows(200,100)
        window.geometry(win_size)
        window.title('Progress functions')
        label = Label(window, text = 'กำลังส่งข้อมูล...', font=FONT)
        label.pack()
        window.update()
        thread = threading.Thread(target = SendMail.finishedmachinedowntime(self, timestamp, machine, inform, problems, action_fix, action_by, notes,location))
        thread.start()
        window.destroy()
        window.update()

    def finishedmachinedowntime(self, timestamp, machine, inform, problems, action_fix, action_by, notes,location):
        subject = 'Machine downtime notification'
        body = (f'''Dear all,
                The problems solved.
                Time occur : {timestamp}
                Machine : {machine}
                Inform by : {inform}
                Problem : {problems}
                Line : {location}
                Solution : {action_fix}
                Solve by : {action_by}
                Notes : {notes}
                Line : {location}
                Status : Closed downtime ''')
        
        to = [self.sendto]
        # to = [self.dow_act]]
        for noti in to:
            try:
                account = create_account(self.username, self.password)
                attachments = []
                send_email(account, subject, body, noti, attachments)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email via outlook [error : {e}]")

# A = SendMail()
# A.finishedmachinedowntime(1,2,3,4,5,6,7,8)





