#coding:utf-8
import os
from datetime import *
from Tkinter import *
import tkMessageBox
import json

from make_word import built_docx
from send_email import SendEmail

config_file = os.path.join(os.getcwd(),"config.ini")

class DailyReport(object):
    def __init__(self, master):
        self.config_data = self.get_config()

        master.title(u"日报")
        master.geometry()

        menubar = Menu(master)
        menubar.add_command(label=u'配置')
        master.config(menu=menubar)

        Label(master, text=u"大数据和平台总体组个人绩效日报")

        #TITLE
        frame = Frame(master, width=380, padx=20, pady=10)
        Label(frame, text=u"大数据和平台总体组个人绩效日报").grid()
        frame.grid(row=0)

        #INFO
        frame_1 = Frame(master, width=380, padx=20, pady=10)

        Label(frame_1, text="FROM:", width=10).grid(row=0, column=0, sticky=W+N)
        Label(frame_1, text="PASSWORD:", width=10).grid(row=0, column=2, sticky=W+N)
        Label(frame_1, text="SERVER:", width=10).grid(row=0, column=4, sticky=W+N)

        self.from_email = StringVar()
        self.password = StringVar()
        self.server = StringVar()

        self.from_email.set(self.config_data['from'])
        self.password.set(self.config_data['password'])
        self.server.set(self.config_data['server'])

        from_entry = Entry(frame_1, textvariable=self.from_email, width=30)
        password_entry = Entry(frame_1, textvariable=self.password, width=30)
        server_entry = Entry(frame_1, textvariable=self.server, width=30)

        from_entry.grid(row=0, column=1)
        password_entry.grid(row=0, column=3)
        server_entry.grid(row=0, column=5)

        Label(frame_1, text="TO:", width=10).grid(row=1, column=0, sticky=W+N)
        Label(frame_1, text="CC:", width=10).grid(row=1, column=3, sticky=W+N)
        # Label(frame_1, text="CC2:", width=10).grid(row=1, column=4, sticky=W+N)

        self.to_email = StringVar()
        self.cc1_email = StringVar()
        # self.cc2_email = StringVar()

        self.to_email.set(self.config_data['to'])
        self.cc1_email.set(self.config_data['cc'])
        # self.cc2_email.set(self.config_data['cc2'])

        to_entry = Entry(frame_1, textvariable=self.to_email, width=42)
        cc1_entry = Entry(frame_1, textvariable=self.cc1_email, width=42)
        # cc2_entry = Entry(frame_1, textvariable=self.cc2_email, width=30)

        to_entry.grid(row=1, column=1, columns=2, sticky=W+N)
        cc1_entry.grid(row=1, column=4,columns=2, sticky=W+N)
        # cc2_entry.grid(row=1, column=5)

        Label(frame_1, text=u"部门", width=10).grid(row=2, column=0, sticky=W+N)
        Label(frame_1, text=u"姓名", width=10).grid(row=2, column=2, sticky=W+N)
        Label(frame_1, text=u"日期", width=10).grid(row=2, column=4, sticky=W+N)

        self.dep_name = StringVar()
        self.name = StringVar()
        self.report_date = StringVar()

        self.dep_name.set(self.config_data['department'])
        self.name.set(self.config_data['name'])
        self.report_date.set(self.today)

        dep_entry = Entry(frame_1, textvariable=self.dep_name, width=30)
        name_entry = Entry(frame_1, textvariable=self.name, width=30)
        date_entry = Entry(frame_1, textvariable=self.report_date, width=30)

        dep_entry.grid(row=2, column=1, sticky=W+N)
        name_entry.grid(row=2, column=3, sticky=W+N)
        date_entry.grid(row=2, column=5, sticky=W+N)

        frame_1.grid(row=1,sticky=W+N)

        #Today tasks info
        frame_2 = Frame(master, width=380, padx=20, pady=10)
        Label(frame_2, text=u"今日任务完成情况").grid()
        frame_2.grid(row=2,sticky=W)
        #Today tasks text area
        frame_2_text = Frame(master, width=380)
        bary2 = Scrollbar(frame_2_text)
        bary2.pack(side=RIGHT,fill=Y)
        self.today_task = Text(frame_2_text, width=150, height=8, padx=2,pady=2)
        self.today_task.pack(side=LEFT,fill=BOTH)
        bary2.config(command=self.today_task.yview)
        self.today_task.config(yscrollcommand=bary2.set)
        frame_2_text.grid(row=3)

        #Unfinished tasks info
        frame_3 = Frame(master, width=380, padx=20, pady=10)
        Label(frame_3, text=u"未完成或遇到问题").grid()
        frame_3.grid(row=4,sticky=W)
        #Unfinished tasks text area
        frame_3_text = Frame(master, width=380)
        bary3 = Scrollbar(frame_3_text)
        bary3.pack(side=RIGHT,fill=Y)
        self.unfinished_task = Text(frame_3_text, width=150, height=8, padx=2,pady=2)
        self.unfinished_task.pack(side=LEFT,fill=BOTH)
        bary3.config(command=self.unfinished_task.yview)
        self.unfinished_task.config(yscrollcommand=bary3.set)
        frame_3_text.grid(row=5)

        #Tomorrow tasks plan
        frame_4 = Frame(master, width=380, padx=20, pady=10)
        Label(frame_4, text=u"明日计划").grid()
        frame_4.grid(row=6,sticky=W)

        #Tomorrow tasks text area
        frame_4_text = Frame(master, width=380)
        bary4 = Scrollbar(frame_4_text)
        bary4.pack(side=RIGHT,fill=Y)
        self.tomorrow_task = Text(frame_4_text, width=150, height=8, padx=2,pady=2)
        self.tomorrow_task.pack(side=LEFT,fill=BOTH)
        bary4.config(command=self.tomorrow_task.yview)
        self.tomorrow_task.config(yscrollcommand=bary4.set)
        frame_4_text.grid(row=7)

        #Thinking
        frame_5 = Frame(master, width=380, padx=20, pady=10)
        Label(frame_5, text=u"收获感悟").grid()
        frame_5.grid(row=8,sticky=W)

        #Thinking text area
        frame_5_text = Frame(master, width=380)
        bary5 = Scrollbar(frame_5_text)
        bary5.pack(side=RIGHT,fill=Y)
        self.think_task = Text(frame_5_text, width=150, height=8, padx=2,pady=2)
        self.think_task.pack(side=LEFT,fill=BOTH)
        bary5.config(command=self.think_task.yview)
        self.think_task.config(yscrollcommand=bary5.set)
        frame_5_text.grid(row=9)


        #Submit
        frame_6 = Frame(master, width=380, padx=20, pady=10)
        submit_btn = Button(frame_6, text=u"提交", command=self.sendmessage, padx=20, pady=10)
        submit_btn.grid(sticky=E)
        frame_6.grid(row=10,sticky=E)

        self.init_data()
    @property
    def today(self):
        tod = datetime.strftime(datetime.now(), "%Y-%m-%d")
        # self.dep_name.set(self.config_data['department'])
        # self.name.set(self.config_data['name'])
        # self.report_date.set(self.tod)
        return tod

    def get_config(self):
        if os.path.exists(config_file):
            with open(config_file,"r") as f:
                configs = f.read()
                configs = json.loads(configs)
                return configs
        configs = {'from':"", 'to':"", 'cc':"", 'department':"", 'name':"", 'password':"", 'server':""}
        with open(config_file,"w") as f:
            f.write(json.dumps(configs))
        return configs

    def check_items(self):
        if len(self.from_email.get()) == 0 or len(self.to_email.get()) == 0 or len(self.dep_name.get()) == 0 or len(self.name.get()) == 0:
            tkMessageBox.showinfo(u"请检查你的设置，发送邮箱，接受邮箱，部门，姓名")
            self.clear()
            self.from_email.focus_set()
            raise
        if len(self.report_date.get()) == 0:
            tkMessageBox.showinfo(u"请检查日报日期")
            self.clear()
            self.report_date.focus_set()
            raise

    def write_config(self, new_config):
        with open(config_file, "w") as f:
            f.write(json.dumps(new_config))
            return True
        return False

    def update_config(self):
        new_config = {}
        new_config['from'] = self.from_email.get()
        new_config['password'] = self.password.get()
        new_config['server'] = self.server.get()
        new_config["to"] = self.to_email.get()
        new_config['cc'] = self.cc1_email.get()
        # new_config["cc1"] = self.cc1_email.get()
        # new_config["cc2"] = self.cc2_email.get()
        new_config["department"] = self.dep_name.get()
        new_config["name"] = self.name.get()
        for k,v in new_config.iteritems():
            if v is None or len(v) == 0:
                return False
        if config_file != self.config_data:
            with open(config_file, "w") as f:
                f.write(json.dumps(new_config))
        return new_config

    # def check_item(self, item, ):
    #     if item is None:
    #         return

    def sendmessage(self):
        print "start send ..."
        self.check_items()
        print "check finished"

        data = self.update_config() if self.update_config() else self.config_data
        data['date'] = self.report_date.get()
        data['today_data'] = self.today_task.get(1.0, END)
        data['unfinished_data'] = self.unfinished_task.get(1.0, END)
        data['tomorrow_data'] = self.tomorrow_task.get(1.0, END)
        data['think_data'] = self.think_task.get(1.0, END)

        daily_docx = built_docx(data)
        if daily_docx:
            data["file_path"] = daily_docx
            se = SendEmail(data)
            se_status = se.send_email()
            if se_status:
                tkMessageBox.showinfo(u"邮件发送成功",daily_docx)
                # self.clear()
            else:
                tkMessageBox.showinfo(u"发送失败，请重试")
                # self.clear()
        else:
            tkMessageBox.showinfo(u"保存失败，请检查。")
#            self.clear()
        # return data


    def init_data(self):
        self.from_email.set("tany@hzhz.co")
        self.password.set("Ty131514")
        self.server.set("smtp.hzhz.co")
        self.to_email.set("tany@hzhz.co")
        self.cc1_email.set("violin748@hotmail.com,tonytan748@gmail.com")
        self.dep_name.set(u"大数据实验室")
        self.name.set(u"谭勇")

        self.today_task.insert(END,u"""
        大数据实验室大数据实验室大数据实验室
        大数据实验室大数据实验室
        大数据实验室大数据实验室大数据实验室""")
        self.unfinished_task.insert(END,u"""
        大数据实验室大数据实验室大数据实验室
        大数据实验室大数据实验室
        大数据实验室大数据实验室大数据实验室""")
        self.tomorrow_task.insert(END,u"""
        大数据实验室大数据实验室大数据实验室
        大数据实验室大数据实验室
        大数据实验室大数据实验室大数据实验室""")
        self.think_task.insert(END,u"""
        大数据实验室大数据实验室大数据实验室
        大数据实验室大数据实验室
        大数据实验室大数据实验室大数据实验室""")


def main():
    root = Tk()
    app = DailyReport(root)
    mainloop()

# main()

if __name__=='__main__':
    main()
