import svgwrite
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta

# === CONFIGURATION ===
num_months = 9  # You can set 3–12
start_month = datetime(2025, 6, 1)

# Layout based on number of months
n_cols = 4 if num_months >= 10 else 3
n_rows = (num_months + n_cols - 1) // n_cols

canvas_width = 1200
canvas_height = 800

# Calendar grid sizing
margin_x = 20
margin_y = 40
spacing_x = 20
spacing_y = 30

calendar_width = (canvas_width - margin_x * 2 - spacing_x * (n_cols - 1)) / n_cols
calendar_height = (canvas_height - margin_y * 2 - spacing_y * (n_rows - 1)) / n_rows

cell_width = calendar_width / 7
cell_height = calendar_height / 8

# Fonts
month_font_size = 14
weekday_font_size = 10
day_font_size = 10
legend_font_size = 10

# Colors
colors = {
    "PPI": "rgb(217, 217, 217)",         # Light grey
    "CPI": "rgb(217, 217, 217)",         # Light grey
    "HLD": "rgb(255, 255, 224)",         # Light yellow
    "WEEKDAY_BG": "rgb(0, 51, 102)",     # Dark blue
    "MONTH_TEXT": "rgb(70, 130, 180)",   # Sky blue
    "DAY_BOX": "white",
}

# Event data
events_by_type = {
    "CPI": ["2025-06-11", "2025-07-15", "2025-08-12", "2025-09-11", "2025-10-15", "2025-11-13", "2025-12-10"],
    "PPI": ["2025-06-12", "2025-07-16", "2025-08-14", "2025-09-10", "2025-10-16", "2025-11-14", "2025-12-11"],
    "HLD": ["2025-07-04", "2025-09-01", "2025-10-13", "2025-11-11", "2025-11-27", "2025-12-25"]
}
event_lookup = {date: label for label, dates in events_by_type.items() for date in dates}

# Create SVG drawing
dwg = svgwrite.Drawing(size=(canvas_width, canvas_height))
dwg.add(dwg.text("Calendar", insert=(canvas_width / 2, 30), text_anchor="middle", font_size=24, font_weight="bold"))

# Draw each month
for idx in range(num_months):
    month_date = start_month + relativedelta(months=idx)
    month = month_date.month
    year = month_date.year
    month_name = month_date.strftime("%B")

    grid_col = idx % n_cols
    grid_row = idx // n_cols

    x_offset = margin_x + grid_col * (calendar_width + spacing_x)
    y_offset = margin_y + grid_row * (calendar_height + spacing_y) + 20

    # Month header
    dwg.add(dwg.text(f"{month_name} {year}", insert=(x_offset, y_offset),
                     font_size=month_font_size, fill=colors["MONTH_TEXT"]))
    y_offset += 15

    # Weekday headers
    weekdays = ["M", "T", "W", "Th", "F", "Sa", "Su"]
    for i, wd in enumerate(weekdays):
        rect = dwg.rect(insert=(x_offset + i * cell_width, y_offset),
                        size=(cell_width, cell_height), fill=colors["WEEKDAY_BG"], stroke="black")
        dwg.add(rect)
        dwg.add(dwg.text(wd,
                         insert=(x_offset + i * cell_width + 2, y_offset + cell_height / 1.5),
                         font_size=weekday_font_size, fill="white"))

    y_offset += cell_height + 2

    # Calendar days
    cal = calendar.Calendar(firstweekday=0)
    days = list(cal.itermonthdays(year, month))
    col, row = 0, 0
    for day in days:
        x = x_offset + col * cell_width
        y = y_offset + row * cell_height
        if day == 0:
            # Empty day box
            dwg.add(dwg.rect(insert=(x, y), size=(cell_width, cell_height),
                             fill=colors["DAY_BOX"], stroke="gray"))
        else:
            date_str = f"{year}-{month:02d}-{day:02d}"
            label = event_lookup.get(date_str)
            fill_color = colors.get(label, colors["DAY_BOX"])

            # Day number box
            dwg.add(dwg.rect(insert=(x, y), size=(cell_width, cell_height),
                             fill=fill_color, stroke="gray"))
            # Day number
            dwg.add(dwg.text(str(day),
                             insert=(x + 2, y + 12),
                             font_size=day_font_size, fill="black"))
            # Event label
            if label:
                dwg.add(dwg.text(label,
                                 insert=(x + 2, y + 24),
                                 font_size=day_font_size - 1, fill="black"))
        col += 1
        if col >= 7:
            col = 0
            row += 1

# Draw Legend
legend_y = canvas_height - 30
legend_x = margin_x
for label in ["PPI", "CPI", "HLD"]:
    dwg.add(dwg.rect(insert=(legend_x, legend_y), size=(10, 10),
                     fill=colors[label], stroke="black"))
    dwg.add(dwg.text(label, insert=(legend_x + 15, legend_y + 9),
                     font_size=legend_font_size, fill="black"))
    legend_x += 80

# Save file
dwg.saveas("calendar_output.svg")
print("✅ Saved to calendar_output.svg")