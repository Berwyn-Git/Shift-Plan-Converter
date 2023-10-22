from tkinter import *
from tkinterdnd2 import *


def path_listbox(event):
    global file_path
    file_path = event.data
    win.destroy()  # close the window after dropping the file


win = TkinterDnD.Tk()
win.title('Drag your file here!')
win.config(bg='gold')

frame = Frame(win)
frame.pack()

listbox = Listbox(
    frame,
    width=50,
    height=15,
    selectmode=SINGLE
)
listbox.pack(fill=X, side=LEFT)
listbox.drop_target_register(DND_FILES)
listbox.dnd_bind('<<Drop>>', path_listbox)

scrolbar = Scrollbar(frame, orient=VERTICAL)
scrolbar.pack(side=RIGHT, fill=Y)

listbox.configure(yscrollcommand=scrolbar.set)
scrolbar.config(command=listbox.yview)

win.mainloop()

print('File path:', file_path)
