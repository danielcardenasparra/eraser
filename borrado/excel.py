import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles.alignment import Alignment

class ExcelImageGUI:
    def __init__(self, master):
        self.master = master
        master.title("Excel Image Creator")
        
        # Create button to save Excel file
        self.save_excel_button = tk.Button(master, text="Save Excel", command=self.save_excel)
        self.save_excel_button.pack()
        
        # Initialize image and Excel file paths
        self.image_path = "telefonica.png"
        self.excel_path = ""
        
    def save_excel(self):
        # Create a new workbook
        wb = Workbook()
        
        # Select the active worksheet
        ws = wb.active
        
        # Load the image into memory
        img = Image(self.image_path)
        
        # Insert the image in the desired cell
        img_cell = ws.cell(row=1, column=1)
        img_cell._value = img
        img_cell.alignment = Alignment(horizontal='center', vertical='center')
        img_anchor = img_cell.coordinate

        # Add the image to the worksheet
        ws.add_image(img, img_anchor)
        
        # Set the width of column A to 450
        ws.column_dimensions['A'].width = 45

        # Save the workbook
        self.excel_path = filedialog.asksaveasfilename(initialdir="/", title="Save Excel", filetypes=(("Excel files", "*.xlsx"),))
        if self.excel_path:
            wb.save(self.excel_path)
            # Show a message indicating the file was saved successfully
            messagebox.showinfo("Success", "Excel file saved successfully.")
        else:
            # If no file is selected, show an error message
            messagebox.showerror("Error", "No file selected.")
        
root = tk.Tk()
gui = ExcelImageGUI(root)
root.mainloop()
