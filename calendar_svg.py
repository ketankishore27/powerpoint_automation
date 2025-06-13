import svgwrite
import calendar
from helpers.helper_files import get_calendar_lookup, check_for_events
from datetime import datetime

def create_calendar_svg(filename):
    # Increased width further to ensure no clipping
    canvas_width = 1180  # Fix for right-side border cutoff
    canvas_height = 750
    dwg = svgwrite.Drawing(filename, profile='tiny', size=(f"{canvas_width}px", f"{canvas_height}px"))

    months = ["March", "April", "May", "June", "July", "August"]
    year = 2025
    colors = {
        "execution": (217, 217, 217),  # Light grey
        "cpi": (255, 255, 224),         # Light yellow
        "ppi": (237, 125, 49),         # Dark orange
        "hld": (198, 224, 180)     # Light green
    }

    cell_width = 53
    cell_height = 40
    x_margin = 25
    y_margin = 50
    top_extra_margin = 30  # For visibility of month titles
    stroke_width = 0.8
    lookup_calendar = get_calendar_lookup()

    
    for idx, month_name in enumerate(months):
        month = idx + 3
        cal = calendar.Calendar(firstweekday=0)
        month_days = list(cal.itermonthdays(year, month))

        grid_col = idx % 3
        grid_row = idx // 3

        x_offset = grid_col * (cell_width * 7 + x_margin)
        y_offset = top_extra_margin + grid_row * (cell_height * 7 + y_margin + 35)

        # Month title
        dwg.add(dwg.text(f"{month_name} {year}",
                         insert=(x_offset, y_offset - 15),
                         font_size="13px", font_weight="bold"))

        # Weekday headers
        weekdays = ["M", "T", "W", "Th", "F", "Sa", "Su"]
        for i, day in enumerate(weekdays):
            dwg.add(dwg.text(day,
                             insert=(x_offset + i * cell_width + 3, y_offset + 10),
                             font_size="9px", font_weight="bold"))

        # Calendar day boxes
        col = 0
        row = 1
        key = None
        for day in month_days:
            if day == 0:
                col += 1
                if col >= 7:
                    col = 0
                    row += 1
                continue

            rect_x = x_offset + col * cell_width
            rect_y = y_offset + row * cell_height

            dwg.add(dwg.rect(insert=(rect_x, rect_y),
                             size=(cell_width, cell_height),
                             fill="#e0f7fa", stroke="gray", stroke_width=stroke_width))

            date_field = f"{day}-{month_name}-{year}"
            date_obj = datetime.strptime(date_field, '%d-%B-%Y')
            formatted_date = date_obj.strftime('%Y-%m-%d')
            key = check_for_events(sample_date=formatted_date, lookup_dictionary=get_calendar_lookup())
            if key:
                dwg.add(dwg.text(str(day) + "\n" + key,
                                 insert=(rect_x + 3, rect_y + 13),
                                 font_size="8px"))
            else:
                dwg.add(dwg.text(str(day),
                                 insert=(rect_x + 3, rect_y + 13),
                                 font_size="8px"))

            col += 1
            if col >= 7:
                col = 0
                row += 1

    dwg.save()
    print(f"✅ Calendar saved to: {filename}")


# Generate SVG
if __name__ == "__main__":
    create_calendar_svg("calendar_march_to_august_2025.svg")