from tkinter import *
from tkinter import messagebox
from main import sheetmetal_flat


root = Tk()
root.title("Macro sheets to DXF for Solidworks")


insert_path = Entry(root, width=100, borderwidth=20)
info_label = Label(root, text="Enter the path for the dxf files")


insert_material = Entry(root, width=100, borderwidth=20)
info_material = Label(root, text="Enter the name of the material")


insert_offset = Entry(root, width=10, borderwidth=20)
info_offset = Label(root, text="Enter an offset value")


relief = BooleanVar()
frame_choice = LabelFrame(root, text="Add automatic reliefs",
                          padx=10, pady=10)
Radiobutton(frame_choice, text="Yes",
            variable=relief, value=True).pack()
Radiobutton(frame_choice, text="No",
            variable=relief, value=False).pack()


thickness = BooleanVar()
frame_thickness = LabelFrame(root, text="Enter the thickness in the name of dxf files",
                          padx=10, pady=10)
Radiobutton(frame_thickness, text="Yes",
            variable=thickness, value=True).pack()
Radiobutton(frame_thickness, text="No",
            variable=thickness, value=False).pack()


insert_path_author = Entry(root, width=100, borderwidth=20)
info_label_author = Label(root, text="The author of this program is Ksawery Fryczynski \n and it is free to use for "
                                     " everyone (even for commercial use) \n who has Solidworks 2019."
                                     , font=("Arial", 25))


def popup_instruction():
    messagebox.showinfo("Instructions",
                        "This is a program to generate DXF flat pattern-sheetmetal files from step assembly \n"
                        "1. Open an assembly file with plates in step format \n" +
                        "2. Select models (you can select all models but the process will be longer) \n" +
                        "3. Select the path, sheet materials and other options in the program window \n" +
                        "4. Click 'Launch macro' \n" +
                        "5. The models that are opened successively appear in the Solidworks window \n" +
                        "6. The program will end with the last open model closed \n" +
                        "7. DXF files will appear in selected path with txt file with undone models \n")


button_instruction = Button(root, text="Instructions and info", command=popup_instruction)

button_macro = Button(root, text="Launch macro",
                      padx=30, pady=50, command=lambda: sheetmetal_flat(path=insert_path.get(),
                                                                        material=insert_material.get(),
                                                                        relief=relief.get(),
                                                                        thickness=thickness.get(),
                                                                        offset=float(insert_offset.get().replace(",",
                                                                                                                 "."))))


insert_path.grid(row=1, column=0)
info_label.grid(row=0, column=0)

insert_material.grid(row=3,column=0)
info_material.grid(row=2, column=0)

insert_offset.grid(row=5,column=0)
info_offset.grid(row=4, column=0)

frame_choice.grid(row=6, column=0)
frame_thickness.grid(row=7, column=0)
button_macro.grid(row=8, column=0)
button_instruction.grid(row=9, column=0)
info_label_author.grid(row=10, column=0)


root.mainloop()
