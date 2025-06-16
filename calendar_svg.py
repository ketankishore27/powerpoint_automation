import svgwrite
import calendar
from helpers.helper_files import get_calendar_lookup, check_for_events
from datetime import datetime

def create_calendar_svg(filename, start_date = "2025-06-01"):
    # Increased width further to ensure no clipping
    canvas_width = 1180  # Fix for right-side border cutoff
    canvas_height = 750
    dwg = svgwrite.Drawing(filename, profile='tiny', size=(f"{canvas_width}px", f"{canvas_height}px"))

    months = ["June", "July", "August", "September", "October", "November"]
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
        month = idx + 6  # June is month 6

        # Create calendar with Monday as the first weekday
        cal = calendar.Calendar(firstweekday=0)
        month_weeks = cal.monthdayscalendar(year, month)

        # Determine grid position
        grid_col = idx % 3
        grid_row = idx // 3

        x_offset = grid_col * (cell_width * 7 + x_margin)
        y_offset = top_extra_margin + grid_row * (cell_height * 7 + y_margin + 35)

        # Add month title
        dwg.add(dwg.text(f"{month_name} {year}",
                         insert=(x_offset, y_offset - 15),
                         font_size="13px", font_weight="bold"))

        # Add weekday headers (aligned with firstweekday=0, i.e., Monday)
        weekdays = ["M", "T", "W", "Th", "F", "Sa", "Su"]
        for i, day in enumerate(weekdays):
            dwg.add(dwg.text(day,
                             insert=(x_offset + i * cell_width + 3, y_offset + 10),
                             font_size="9px", font_weight="bold"))

        ## Replace code start
        for row_idx, week in enumerate(month_weeks):
            for col, day in enumerate(week):
                if day == 0:
                    continue  # skip padding days outside this month

                # Calculate position
                rect_x = x_offset + col * cell_width
                rect_y = y_offset + (row_idx + 1) * cell_height  # +1 for header row

                # Draw the cell rectangle
                dwg.add(dwg.rect(insert=(rect_x, rect_y),
                                 size=(cell_width, cell_height),
                                 fill="#e0f7fa", stroke="gray", stroke_width=stroke_width))

                # Format date and check for event labels
                date_field = f"{day}-{month_name}-{year}"
                try:
                    date_obj = datetime.strptime(date_field, '%d-%B-%Y')
                except ValueError:
                    continue

                formatted_date = date_obj.strftime('%Y-%m-%d')
                key = check_for_events(sample_date=formatted_date, lookup_dictionary=lookup_calendar)

                # Prepare label (day + optional event)
                label = str(day) + ("\n" + key if key else "")

                # Draw the text
                dwg.add(dwg.text(label,
                                 insert=(rect_x + 3, rect_y + 13),
                                 font_size="8px"))
            
        ## Replace code end

    dwg.save()
    print(f"✅ Calendar saved to: {filename}")


# Generate SVG
if __name__ == "__main__":
    create_calendar_svg("calendar_march_to_august_2025.svg")