import os
#import win32com.client as win32
import pandas as pd
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta
from data.data_mockers import get_data_captiatalization
import pandas as pd

def create_ppt_with_template(template_path):#, output_path, df):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"❌ Template not found at {template_path}")

    # Launch PowerPoint
    ppt_app = win32.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True

    # Open the template as the base presentation
    presentation_sample = ppt_app.Presentations.Open(template_path)
    return presentation_sample

def create_current_captiatlization_helper(presentation_sample):

    output_folder = os.path.join(os.getcwd(), "output_path")
    slide = presentation_sample.Slides.Add(presentation_sample.Slides.Count + 1, 1)
    title_shape = slide.Shapes.Title
    title_shape.TextFrame.TextRange.Text = "Current Capitalization"
    title_shape.Left = 10    # distance from the left (in points)
    title_shape.Top = 30     # distance from the top (in points)
    title_shape.Width = 400  # optional: adjust width
    title_shape.Height = 30  # optional: adjust height
    
    highlight_rows = {"First Lien Debt", "Net First Lien Debt", "Total Debt", "Total Net Debt"}
    footer_row = "LTM Run-Rate Adj. EBITDA"
    data = get_data_captiatalization()
    
    ## Getting number of rows and cols 
    rows = len(data)
    cols = len(data[0])
    
    left = 20     # points
    top = 70
    width = 900
    height = 300
    table_shape = slide.Shapes.AddTable(rows, cols, left, top, width, height)
    table = table_shape.Table

    # Filling Data here
    for i, row_data in enumerate(data):
        for j, cell_text in enumerate(row_data):
            cell = table.Cell(i + 1, j + 1)
            cell.Shape.TextFrame.TextRange.Text = cell_text
            cell.Shape.TextFrame.TextRange.Font.Size = 8
            cell.Shape.TextFrame.TextRange.Font.Bold = True if i == 0 or row_data[0] in highlight_rows or row_data[0] == footer_row else False
            cell.Shape.TextFrame.VerticalAnchor = 3  # center vertically
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 if j != 0 else 1  # right or left align
    
        # Row shading
        if row_data[0] in highlight_rows:
            for j in range(cols):
                table.Cell(i + 1, j + 1).Shape.Fill.ForeColor.RGB = 0xCCE5FF  # light blue
        elif row_data[0] == footer_row:
            for j in range(cols):
                table.Cell(i + 1, j + 1).Shape.Fill.ForeColor.RGB = 0xD9CCE3  # light purple

    output_file_path = os.path.join(output_folder, "captiatlisation_sheet.pptx")
    presentation_sample.SaveAs(output_file_path)
    print("✅ Capitalisation Slides Done.")
    return output_file_path


def get_calendar_lookup():
    lookup_df = pd.read_excel("data/events_data.xlsx")
    lookup_df["Date"] = lookup_df["Date"].apply(lambda x: str(datetime.date(x)))
    lookup_dict = {}
    for date, event in lookup_df.values:
        try:
            lookup_dict[event].append(date)
        except Exception as e:
            lookup_dict[event] = []
    return lookup_dict


def check_for_events(sample_date, lookup_dictionary):
    for key in lookup_dictionary.keys():
        if sample_date in lookup_dictionary[key]:
            return key
    return None

# Start date
def get_calendar_slide_helper(presentation_sample, date_str = '2025-04-01'):

    output_folder = os.path.join(os.getcwd(), "output_path")
    start_date = datetime.strptime(date_str, '%Y-%m-%d')
    
    months = []
    for i in range(6):
        date = start_date + relativedelta(months=i)
        months.append({
            'month_name': date.strftime('%B'),     # e.g. 'April'
            'month_number': date.strftime('%m'),   # e.g. '04'
            'year': date.strftime('%Y')            # e.g. '2025'
        })
    
    colors = {
        "execution": (217, 217, 217),  # Light grey
        "cpi": (255, 255, 224),         # Light yellow
        "ppi": (237, 125, 49),         # Dark orange
        "hld": (198, 224, 180)     # Light green
    }
    lookup_calendar = get_calendar_lookup()
    # Initialize PowerPoint
    ppt = win32.Dispatch("PowerPoint.Application")
    ppt.Visible = True
    slide = presentation_sample.Slides.Add(presentation_sample.Slides.Count + 1, 12)
    
    # Title
    title_box = slide.Shapes.AddTextbox(1, 20, 10, 700, 50)
    title_box.TextFrame.TextRange.Text = "*EXAMPLE* - Overview of Execution Windows"
    title_box.TextFrame.TextRange.Font.Size = 24
    
    # Grid layout parameters
    cal_per_row = 3
    calendar_width = 200
    calendar_height = 160
    cell_spacing = 1
    cell_w = (calendar_width - (6 * cell_spacing)) / 7  # 7 columns
    cell_h = (calendar_height - 25 - (5 * cell_spacing)) / 6  # 6 rows
    
    margin_left = 50
    margin_top = 70
    h_spacing = 60
    v_spacing = 50  # 🔻 Reduced from 70 to make rows closer
    
    # RGB converter
    def rgb(r, g, b):
        return r + (g << 8) + (b << 16)
    
    # Create calendars
    for i, month in enumerate(months):
        month_idx = list(calendar.month_name).index(month['month_name'])
        month_matrix = calendar.monthcalendar(int(month['year']), month_idx)
    
        row = i // cal_per_row
        col = i % cal_per_row
    
        left = margin_left + col * (calendar_width + h_spacing)
        top = margin_top + row * (calendar_height + v_spacing)
    
        # Month Title
        title = slide.Shapes.AddTextbox(1, left, top, calendar_width, 20)
        title.TextFrame.TextRange.Text = f"{month['month_name']} {month['year']}"
        title.TextFrame.TextRange.Font.Size = 14
        title.TextFrame.TextRange.Font.Bold = True
    
        # Day cells
        for week_idx, week in enumerate(month_matrix):
            for day_idx, day in enumerate(week):
                if day == 0:
                    continue
                    
                key = None
                shape = slide.Shapes.AddTextbox(
                    1,
                    left + day_idx * (cell_w + cell_spacing),
                    top + 25 + week_idx * (cell_h + cell_spacing),
                    cell_w,
                    cell_h
                )
                
                date_looped = str(datetime.date(datetime.strptime(month['year'] + '-' + month['month_number'] + '-' + str(day), '%Y-%m-%d')))
                key = check_for_events(sample_date=date_looped, lookup_dictionary=get_calendar_lookup())
                if key:
                    shape.TextFrame.TextRange.Text = str(day) + f"\n{key}"
                    shape.Fill.ForeColor.RGB = rgb(*colors[key.lower()])
                    shape.TextFrame.TextRange.Font.Bold = True
                else:
                    shape.TextFrame.TextRange.Text = str(day) + f"\n "
                    shape.Fill.ForeColor.RGB = rgb(*colors["execution"])
                    
                shape.TextFrame.TextRange.Font.Size = 6
    
                # Bottom-right align
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = 3  # Right
                shape.TextFrame.VerticalAnchor = 3  # Bottom
    
                # Fill with Execution Window color
                shape.Line.Visible = False
    
    # Add legend boxes
    legend_items = [
        ("CPI", colors["cpi"]),
        ("PPI", colors["ppi"]),
        ("Holiday", colors["hld"]),
    ]
    
    legend_top = margin_top + 2 * (calendar_height + v_spacing) + 10
    legend_left = margin_left
    
    for i, (label, color_rgb) in enumerate(legend_items):
        box = slide.Shapes.AddShape(1, legend_left + i * 180, legend_top, 20, 20)
        box.Fill.ForeColor.RGB = rgb(*color_rgb)
        box.Line.Visible = False
    
        label_box = slide.Shapes.AddTextbox(1, legend_left + i * 180 + 25, legend_top, 150, 20)
        label_box.TextFrame.TextRange.Text = label
        label_box.TextFrame.TextRange.Font.Size = 12

    output_calendar_path = os.path.join(output_folder, "calendar_sheet.pptx")
    presentation_sample.SaveAs(output_calendar_path)
    print("✅ Calendar Slides Done.")
    return output_calendar_path
