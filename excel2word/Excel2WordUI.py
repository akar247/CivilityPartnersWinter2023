from tkinter import *
from tkinter import filedialog
from tkinter import font
from tkinter.ttk import Progressbar
from tkinter.ttk import Scrollbar


#main window
window = Tk()
window.title("Revenue Categorization")
window.geometry('600x300')
window.config(bg='#ffffff')

bg_color = 'white'


# frames
frame1 = Frame(bg='#3b84f8') # holds the title
frame1.pack(fill='x')
frame6 = Frame(bg='#ffffff') # holds the interactive message
frame6.pack()
frame2 = Frame(bg='#ffffff') # holds the choose file, preview, and start buttons
frame2.pack()
frame3 = Frame(bg='#ffffff') # holds the progress bar
frame3.pack(fill='x')

# handler for 'Choose File' button click event
def open_xfile():
    global xfile
    xfile = filedialog.askopenfilename(title='open a file',filetypes=(('excel files','*.xls'),('excel files','*.xlsx')))
    lbl_filename.config(text=filename,fg='grey')


#BELOW IS ALL UI
# create widgets (labels, buttons, progress bar, scroll bar)
lbl_title = Label(master=frame1,font=('Arial',25,'bold'),text='Revenue Categorization Script',fg='white',bg='#3b84f8')
lbl_filename = Label(master=frame2,text='No File Chosen',fg='grey',width=20,anchor='w')
lbl_message = Label(master=frame6,text='*Please upload a file',bg='white', fg='#3b84f8')

btn_choosefile = Button(master=frame2,text='Choose File',command=open_xfile, fg = "white", font = "Future 10", bg="#3b84f8")

progress_bar = Progressbar(master=frame3, orient=HORIZONTAL, length=500, mode='determinate')

# pack widgets
lbl_title.pack(padx=20,pady=20)
lbl_filename.grid(sticky='w',row=0,column=1,padx=10,pady=10)
lbl_message.pack(pady=5)

btn_choosefile.grid(sticky='w',row=0,column=0,padx=10,pady=10)

progress_bar.pack(padx=10,pady=20,expand=True)

# execute
window.mainloop() 
