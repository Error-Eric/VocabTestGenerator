import vocabtest
import tkinter as tk
from tkinter import Menu, filedialog, messagebox

class CollapsibleFrame(tk.Frame):
    def __init__(self, parent, title):
        super().__init__(parent)
        
        self.title, self.is_collapsed = title, True
        
        # Create a label for the title
        self.title_label = tk.Label(self, text=self.title, bg='lightgray', relief=tk.RAISED)
        self.title_label.pack(fill=tk.X)
        
        # Create a frame for the content
        self.content_frame = tk.Frame(self)
        
        # Add a button to toggle the content
        self.title_label.bind("<Button-1>", self.toggle_content)
        
    def toggle_content(self, event):
        """Toggle the visibility of the content frame."""
        if self.is_collapsed:
            self.content_frame.pack(fill=tk.BOTH, expand=True)
            self.is_collapsed = False
        else:
            self.content_frame.pack_forget()
            self.is_collapsed = True

    def add_content(self, widget):
        """Add a widget to the content frame."""
        widget.pack(padx=10, pady=5)

def on_exit():
    """Exit the application."""
    root.quit()

def opentable():
    global curtest
    filename = filedialog.askopenfilename(
                                          title="Select an Excel File",
                                          filetypes=(("Excel files",
                                                       "*.xls;*.xlsx;*.xlsm"),
                                                      ))
    print(f"{filename} has been selected.")
    curtest = vocabtest.vocabtest(filename)
    messagebox.showinfo("File read", f"""Excel file {filename} has been loaded, 
                        {curtest.numunit()} units and {curtest.numword()} words 
                        has been loaded.""")

    for i, ux in enumerate(curtest.unitdicts):
        if len(ux) != 0: 
            print(f'Unit {i} detected.')
            unitcontent = CollapsibleFrame(scrollable_frame, f'Unit {i}') 
            unitcontent.pack(fill=tk.X, padx=5, pady=5)  # Pack the CollapsibleFrame
            
            # Create a label to show str(ux) with line wrapping
            unit_label = tk.Label(unitcontent.content_frame, text=str(ux), wraplength=600, justify="left")
            unitcontent.add_content(unit_label)  # Add the label to the content frame

def open_gen_dialog():
    """Open a dialog for generate a test."""
    dialog = tk.Toplevel(root)
    dialog.title("Generate vocab test")

    # Create labels and entry fields
    tk.Label(dialog, text="The units you want to test:").pack(padx=10, pady=5)
    string_entry = tk.Entry(dialog)
    string_entry.pack(padx=10, pady=5)

    tk.Label(dialog, text="The number of words in the test:").pack(padx=10, pady=5)
    number_entry = tk.Entry(dialog)
    number_entry.pack(padx=10, pady=5)

    tk.Label(dialog, text="Name of the test:").pack(padx=10, pady=5)
    name_entry = tk.Entry(dialog)
    name_entry.pack(padx=10, pady=5)

    def confirm():
        rawunit, unit = string_entry.get().replace('，', ',').split(','), []
        for rx in rawunit:
            try: unit.append(int(rx))
            except: 
                try: 
                    ul, ur = map(int, rx.split('-'))
                    if ul + 40 > ur and ul <= ur:
                        for i in range(ul, ur+1): unit.append(i)
                    else : raise KeyError("Bad range")
                except: pass # raise error window
        unit.sort()
        totnum, testname = int(number_entry.get()), name_entry.get()
        # You can add validation here if needed
        print(f"Units: {unit}, Number: {totnum}")
        curtest.gentest(units = unit, num = totnum, title= testname)
        dialog.destroy()  # Close the dialog

    def cancel():
        dialog.destroy()  # Close the dialog without taking action

    # Create buttons
    confirm_button = tk.Button(dialog, text="Confirm", command=confirm)
    confirm_button.pack(padx=10, pady=5)
    
    cancel_button = tk.Button(dialog, text="Cancel", command=cancel)
    cancel_button.pack(padx=10, pady=5)

def open_save_dialog():
    if curtest == None:
        print("Nothing opened.")
        return
    filename = filedialog.asksaveasfilename(
        confirmoverwrite= True, 
        filetypes= (("Excel files","*.xlsx"),),                                         
        defaultextension= ".xlsx")
    print(f"saving to {filename}")
    curtest.puttest(True, filename)

# Create the main application window
root = tk.Tk()
root.title("My Tkinter Application")
root.geometry("720x480")  # Set the window size to 720x480

# Create a menu bar
menu_bar = Menu(root)

# Add the commands to the menu bar
menu_bar.add_command(label="Open", command=opentable) 
menu_bar.add_command(label="Generate", command= open_gen_dialog)
menu_bar.add_command(label="Save", command=open_save_dialog)
menu_bar.add_command(label="Exit", command=on_exit)

# Configure the menu bar
root.config(menu=menu_bar)

# Create a canvas for scrolling
canvas = tk.Canvas(root)
scrollable_frame = tk.Frame(canvas)

# Create a scrollbar
scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

# Pack the scrollbar and canvas
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Create a window in the canvas
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

# Update the scroll region when the frame is resized
def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

scrollable_frame.bind("<Configure>", on_frame_configure)

# Bind mouse wheel scrolling
def on_mouse_wheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind_all("<MouseWheel>", on_mouse_wheel)  # For Windows and Mac
canvas.bind_all("<Button-4>", lambda event: canvas.yview_scroll(-1, "units"))  # For Linux
canvas.bind_all("<Button-5>", lambda event: canvas.yview_scroll(1, "units"))  # For Linux

curtest = None

# Start the Tkinter main loop
root.mainloop()
