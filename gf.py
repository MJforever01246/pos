import tkinter as tk
from tkinter import ttk
from tkinter import*


root = tk.Tk()

def validate(newtext):
    print('validate: {}'.format(newtext))
    return True
vcmd = root.register(validate)

def key(event):
    print('key: {}'.format(event.char))

def var(*args):
    print('var: {} (args {})'.format(svar.get(), args))
svar = tk.StringVar()
svar.trace('w', var)

entry = tk.Entry(root,
                 textvariable=svar,
                 validate="key", validatecommand=(vcmd, '%P'))
entry.bind('<Key>', key)
entry.pack()
root.mainloop()
# class App:
#     def __init__(self):
#         self.root = tk.Tk()
#         self.tree = ttk.Treeview()
#         self.tree.pack()
#         for i in range(10):
#             self.tree.insert("", "end", text="Item %s" % i)
#         self.tree.bind("<Double-1>", self.OnDoubleClick)
#         self.root.mainloop()
#
#     def OnDoubleClick(self, event):
#         print(self.tree.item(self.tree.get_children()[0]))
#
# if __name__ == "__main__":
#     app = App()