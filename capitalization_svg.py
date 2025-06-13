from svgwrite import Drawing
from svgwrite.container import Group
from svgwrite.text import Text
from data.data_mockers import get_data_captiatalization

def add_Captialization_svg(filename):
    # Table data as rows of values
    # Table data as rows of values
    table_data = get_data_captiatalization()

    # Define styling
    cell_width = 100
    cell_height = 40
    font_size = 13
    title_font_size = 30
    num_cols = len(table_data[0])
    num_rows = len(table_data)

    # Canvas size
    width = cell_width * num_cols + 50
    height = cell_height * (num_rows + 2)  # +2 for spacing and title

    # Create drawing
    
    dwg = Drawing(filename, size=(width, height))
    group = Group()

    # Draw title
    group.add(Text("Current Capitalization", insert=(20, title_font_size + 20), font_size=title_font_size, font_weight="bold"))

    # Colors
    header_color = "#009CA6"
    even_row_color = "#FFFFFF"
    highlight_color = "#FDE0C2"
    footer_color = "#E4C4DC"
    text_color = "#000"

    # Draw table
    start_y = title_font_size + 50
    for row_idx, row in enumerate(table_data):
        y = start_y + row_idx * cell_height
        for col_idx, cell in enumerate(row):
            x = 20 + col_idx * cell_width

            # Determine background color
            if row_idx == 0:
                fill = header_color
                font_weight = "bold"
                font_fill = "white"
            elif row_idx in [6, 7, 11, 12]:
                fill = highlight_color
                font_weight = "bold"
                font_fill = text_color
            elif row_idx == 13:
                fill = footer_color
                font_weight = "bold"
                font_fill = text_color
            else:
                fill = even_row_color
                font_weight = "normal"
                font_fill = text_color

            # Draw cell background
            group.add(dwg.rect(insert=(x, y), size=(cell_width, cell_height), fill=fill, stroke="black", stroke_width=1))

            # Draw cell text
            if cell:
                group.add(Text(
                    cell,
                    insert=(x + 5, y + cell_height / 2 + font_size / 2.5),
                    font_size=font_size,
                    fill=font_fill,
                    font_weight=font_weight
                ))

    dwg.add(group)
    dwg.save()

if __name__ == "__main__":
    add_Captialization_svg(filename="capitalization_table.svg")
