import tkinter as tk
from tkinter import filedialog
import pandas as pd

root = tk.Tk()

canvas1 = tk.Canvas(root, width=300, height=300, bg='lightsteelblue')
canvas1.pack()


def getPhoto():
    filename = filedialog.askopenfilename()
    photo = tk.PhotoImage(file=filename)
    print(filename)


browseButton_Photo = tk.Button(text='Import Photo', command=getPhoto, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_Photo)
canvas1.configure(image=photo)

root.mainloop()
