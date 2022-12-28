import tkinter
main_window = tkinter.Tk()

label1 = tkinter.Label(main_window, text="Save")
label2 = tkinter.Label(main_window, text="Open")
tombol1 = tkinter.Button(main_window, text='Tombol1')
tombol2 = tkinter.Button(main_window, text='Tombol2')

label1.pack()
label2.pack()
tombol1.pack()
tombol2.pack()

main_window.mainloop()