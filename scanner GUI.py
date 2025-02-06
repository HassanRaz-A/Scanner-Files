import os
import win32com.client
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class ExcelToPPTConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PowerPoint Converter")
        self.root.geometry("600x400")
        
        self.create_widgets()
        self.image_pairs = {
            "Best Pilot and Beam": [
                ("Picture 6", "Picture 7"),
                ("Picture 8", "Picture 9"),
                ("Picture 10", "Picture 11"),
                ("Picture 2", "Picture 3"),
            ],
            "1_LIST": [
                ("Picture 6", "Picture 7"),
                ("Picture 8", "Picture 9"),
                ("Picture 10", "Picture 11"),
                ("Picture 4", "Picture 5"),
            ]
        }
        self.SCALE_FACTORS = [
            {"height": 1.30, "width": 1.30},
            {"height": 1.15, "width": 1.15},
        ]

    def create_widgets(self):
        # File Selection
        ttk.Label(self.root, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.excel_entry = ttk.Entry(self.root, width=50)
        self.excel_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.root, text="Browse", command=self.browse_excel).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(self.root, text="PPT Template:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.ppt_entry = ttk.Entry(self.root, width=50)
        self.ppt_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.root, text="Browse", command=self.browse_ppt).grid(row=1, column=2, padx=5, pady=5)

        # Slide Numbers
        ttk.Label(self.root, text="Starting Slide for 'Best Pilot and Beam':").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.slide_entry1 = ttk.Entry(self.root, width=10)
        self.slide_entry1.grid(row=2, column=1, padx=5, pady=5, sticky='w')

        ttk.Label(self.root, text="Starting Slide for '1_LIST':").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        self.slide_entry2 = ttk.Entry(self.root, width=10)
        self.slide_entry2.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        # Run Button
        self.run_btn = ttk.Button(self.root, text="Run Conversion", command=self.run_conversion)
        self.run_btn.grid(row=4, column=1, pady=20)

        # Status Log
        self.status = tk.Text(self.root, height=8, width=70)
        self.status.grid(row=5, column=0, columnspan=3, padx=5, pady=5)
        self.status.insert(tk.END, "Status messages will appear here...")
        self.status.config(state=tk.DISABLED)

    def browse_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.excel_entry.delete(0, tk.END)
        self.excel_entry.insert(0, filepath)

    def browse_ppt(self):
        filepath = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx")])
        self.ppt_entry.delete(0, tk.END)
        self.ppt_entry.insert(0, filepath)

    def log_message(self, message):
        self.status.config(state=tk.NORMAL)
        self.status.insert(tk.END, message + "\n")
        self.status.see(tk.END)
        self.status.config(state=tk.DISABLED)
        self.root.update()

    def run_conversion(self):
        excel_path = self.excel_entry.get()
        ppt_path = self.ppt_entry.get()
        slide1 = self.slide_entry1.get()
        slide2 = self.slide_entry2.get()

        try:
            slide_numbers = {
                "Best Pilot and Beam": int(slide1),
                "1_LIST": int(slide2)
            }
        except ValueError:
            messagebox.showerror("Error", "Invalid slide numbers")
            return

        if not all([excel_path, ppt_path]):
            messagebox.showerror("Error", "Please select both Excel and PowerPoint files")
            return

        try:
            self.log_message("Starting conversion process...")
            
            # Verify files exist
            if not os.path.exists(excel_path):
                raise FileNotFoundError(f"Excel file not found: {excel_path}")
            if not os.path.exists(ppt_path):
                raise FileNotFoundError(f"PowerPoint template not found: {ppt_path}")

            # Open Excel
            try:
                ExcelApp = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                ExcelApp = win32com.client.Dispatch("Excel.Application")
            ExcelApp.Visible = True

            # Open Excel workbook
            try:
                xlWorkbook = ExcelApp.Workbooks.Open(excel_path)
            except Exception as e:
                raise Exception(f"Failed to open Excel workbook: {e}")

            # Open PowerPoint
            PPTApp = win32com.client.Dispatch("PowerPoint.Application")
            PPTApp.Visible = True
            PPTPresentation = PPTApp.Presentations.Open(ppt_path)

            # Process sheets
            for sheet_name in ["Best Pilot and Beam", "1_LIST"]:
                self.log_message(f"\nProcessing sheet: {sheet_name}")
                try:
                    xlWorksheet = xlWorkbook.Worksheets(sheet_name)
                except Exception:
                    self.log_message(f"Sheet '{sheet_name}' not found, skipping.")
                    continue

                xlShapes = xlWorksheet.Shapes
                pairs = self.image_pairs[sheet_name]
                slide_number = slide_numbers[sheet_name]

                for pair in pairs:
                    if slide_number > len(PPTPresentation.Slides):
                        PPTPresentation.Slides.Add(slide_number, 1)

                    PPTSlide = PPTPresentation.Slides(slide_number)
                    pasted_shapes = []

                    for index, img_name in enumerate(pair):
                        for xlShape in xlShapes:
                            if xlShape.Type == 13 and xlShape.Name == img_name:
                                try:
                                    xlShape.Copy()
                                    shape = PPTSlide.Shapes.PasteSpecial()
                                    shape_item = shape.Item(1)

                                    # Scale image
                                    shape_item.Width *= self.SCALE_FACTORS[index]["width"]
                                    shape_item.Height *= self.SCALE_FACTORS[index]["height"]

                                    pasted_shapes.append(shape_item)
                                    self.log_message(f"Copied {xlShape.Name} to slide {slide_number}")

                                except Exception as e:
                                    self.log_message(f"Error copying {xlShape.Name}: {str(e)}")

                    # Position images
                    if len(pasted_shapes) == 2:
                        slide_width = PPTSlide.Master.Width
                        slide_height = PPTSlide.Master.Height

                        pasted_shapes[0].Left = (slide_width - pasted_shapes[0].Width) / 2
                        pasted_shapes[0].Top = (slide_height - pasted_shapes[0].Height) / 2

                        pasted_shapes[1].Left = pasted_shapes[0].Left + pasted_shapes[0].Width
                        pasted_shapes[1].Top = pasted_shapes[0].Top

                    slide_number += 1

            # Save and close
            PPTPresentation.SaveAs(ppt_path)
            self.log_message("\nConversion completed successfully!")
            xlWorkbook.Close(SaveChanges=False)
            ExcelApp.Quit()
            PPTApp.Quit()

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.log_message(f"\nError occurred: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToPPTConverter(root)
    root.mainloop()
