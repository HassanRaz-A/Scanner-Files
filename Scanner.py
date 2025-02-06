import os
import win32com.client

file_path = r"C:\Users\user\DTS Dropbox\Ehsan Nawaz\Old Dropbox files\Post Processing\T-Mobile\Yale Schwartzman\Raw Reports\Scanner\Mezz\Scanner_Mezz_23.NR_TopN_Beam_CellId_Best_FR1 FDD n71 DL_5G_NR_CH124570_INDOOR.xlsx"
template_path = r"C:\Users\user\DTS Dropbox\Ehsan Nawaz\Old Dropbox files\Post Processing\T-Mobile\Yale Schwartzman\Summary Reports\Mezz\Scanner.pptx"
# Verify if files exist
if not os.path.exists(file_path):
    raise FileNotFoundError(f"Excel file not found: {file_path}")
if not os.path.exists(template_path):
    raise FileNotFoundError(f"PowerPoint template not found: {template_path}")

# Open Excel
try:
    ExcelApp = win32com.client.GetActiveObject("Excel.Application")
except Exception:
    ExcelApp = win32com.client.Dispatch("Excel.Application")
ExcelApp.Visible = True

# Open Excel workbook
try:
    xlWorkbook = ExcelApp.Workbooks.Open(file_path)
except Exception as e:
    raise Exception(f"Failed to open Excel workbook: {e}")

# Open PowerPoint
PPTApp = win32com.client.Dispatch("PowerPoint.Application")
PPTApp.Visible = True
PPTPresentation = PPTApp.Presentations.Open(template_path)

# Image pairs for specific sheets
image_pairs = {
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

# Define scaling factors
SCALE_FACTORS = [
    {"height": 1.30, "width": 1.30},  # Image 1 scaling (130%)
    {"height": 1.15, "width": 1.15},  # Image 2 scaling (125%)
]

# Process sheets
for sheet_name in ["Best Pilot and Beam", "1_LIST"]:
    try:
        xlWorksheet = xlWorkbook.Worksheets(sheet_name)
    except Exception:
        print(f"Sheet '{sheet_name}' not found, skipping.")
        continue

    xlShapes = xlWorksheet.Shapes
    pairs = image_pairs[sheet_name]

    # Ask for the first slide number
    slide_number = int(input(f"Enter the starting slide number for '{sheet_name}': "))

    for pair in pairs:
        if slide_number > len(PPTPresentation.Slides):
            PPTPresentation.Slides.Add(slide_number, 1)  # Add new slide if needed

        PPTSlide = PPTPresentation.Slides(slide_number)
        pasted_shapes = []

        for index, img_name in enumerate(pair):
            for xlShape in xlShapes:
                if xlShape.Type == 13 and xlShape.Name == img_name:  # Type 13 = Picture
                    try:
                        xlShape.Copy()
                        shape = PPTSlide.Shapes.PasteSpecial()
                        shape_item = shape.Item(1)

                        # Scale image based on predefined factors
                        shape_item.Width *= SCALE_FACTORS[index]["width"]
                        shape_item.Height *= SCALE_FACTORS[index]["height"]

                        pasted_shapes.append(shape_item)  # Store for positioning

                        print(f"Copied {xlShape.Name} from '{sheet_name}' to slide {slide_number}")

                    except Exception as e:
                        print(f"Failed to copy {xlShape.Name} from '{sheet_name}': {e}")

        # Align images as required
        if len(pasted_shapes) == 2:
            slide_width = PPTSlide.Master.Width
            slide_height = PPTSlide.Master.Height

            # Center the first image
            pasted_shapes[0].Left = (slide_width - pasted_shapes[0].Width) / 2
            pasted_shapes[0].Top = (slide_height - pasted_shapes[0].Height) / 2

            # Position the second image **to the right** of the first image (parallel)
            pasted_shapes[1].Left = pasted_shapes[0].Left + pasted_shapes[0].Width + 0  # Adding gap of 20 units
            pasted_shapes[1].Top = pasted_shapes[0].Top  # Keep it parallel

        slide_number += 1  # Move to next slide

# Save PowerPoint directly to the template file (overwrite the template)
PPTPresentation.SaveAs(template_path)
print(f"Presentation updated and saved: {template_path}")

# Close Excel & PowerPoint
xlWorkbook.Close(SaveChanges=False)
ExcelApp.Quit()
PPTApp.Quit()
