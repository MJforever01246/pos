import tkinter as tk
from tkinter import*
from CustomerManagement import CustomerManagement
from EmployeesManagement import EmployeesManagement
from Mainsession import MainSession,ChangePassword
from Inventory import Inventory
from Report import Report
class Dashboard(tk.Tk):
    def __init__(self,account,position,username):
        tk.Tk.__init__(self)
        # create window
        self.geometry("1150x600+400+400")
        self.title("Bảng quản lí")
        self.resizable(0, 0)
        back = Frame(master=self, bg='white')
        back.pack_propagate(0)
        back.pack(expand=1)
        # button
        canvas = Canvas(self, width=1150, height=600, bg="#000000")
        bg = Label(self)
        bg.place(relx=0, rely=0, width=1150, height=600)
        bg_img = PhotoImage(master = canvas, width = 1150, height = 600,file="dashboard.png")
        bg.configure(image=bg_img)

        log_out_img = PhotoImage(master=canvas, file="logout_dashboard.png")
        log_out_button = Button(image=log_out_img, command=self.Log_out, relief="flat", overrelief="flat",
                                activebackground="#ffffff", cursor="hand2",
                                foreground="#ffffff", background="#ffffff", borderwidth="0")
        log_out_button.place(x=986, y=13)

        customer_img = PhotoImage(master=canvas, file="customer_dashboard.png")
        customer_button = Button(image=customer_img, command=CustomerManagement, relief="flat", overrelief="flat",
                                activebackground="#ffffff", cursor="hand2",
                                foreground="#ffffff", background="#ffffff", borderwidth="0")
        customer_button.place(x=111, y=185)

        report_img = PhotoImage(master=canvas, file="report_dashboard.png")
        report_button = Button(image=report_img,command=Report, relief="flat", overrelief="flat",
                                activebackground="#ffffff", cursor="hand2",
                                foreground="#ffffff", background="#ffffff", borderwidth="0")
        report_button.place(x=815, y=183)

        employee_img = PhotoImage(master=canvas, file="user_dashboard.png")
        employee_button = Button(image=employee_img, command=EmployeesManagement, relief="flat", overrelief="flat",
                                activebackground="#ffffff", cursor="hand2",
                                foreground="#ffffff", background="#ffffff", borderwidth="0")
        employee_button.place(x=111, y=356)

        pos_img = PhotoImage(master=canvas, file="pos_dashboard.png")
        pos_button = Button(master=self,image=pos_img, command=lambda:MainSession(account,position,username), relief="flat", overrelief="flat",
                                activebackground="#ffffff", cursor="hand2",
                                foreground="#ffffff", background="#ffffff", borderwidth="0")
        pos_button.place(x=815, y=356)

        product_img = PhotoImage(master=canvas, file="product_dashboard.png")
        product_button = Button(image=product_img, command=Inventory, relief="flat", overrelief="flat",
                                activebackground="#ffffff", cursor="hand2",
                                foreground="#ffffff", background="#ffffff", borderwidth="0")
        product_button.place(x=111, y=531)

        change_password_img = PhotoImage(master=canvas, file="changepassword_dashboard.png")
        change_password_button = Button(image=change_password_img, command=lambda:ChangePassword(username=username), relief="flat", overrelief="flat",
                                activebackground="#ffffff", cursor="hand2",
                                foreground="#ffffff", background="#ffffff", borderwidth="0")
        change_password_button.place(x=815, y=531)
        self.mainloop()
    def Log_out(self):
        self.destroy()