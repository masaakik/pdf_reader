import os
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont

def create_inspection_stamp(name_top="河", name_bottom="本", output_path="temp_stamp.png") -> str:
    """電子印鑑（デーツスタンプ）を自動生成する"""
    size = 400
    padding = 10
    color_vermilion = "#FF4500"
    font_path = "C:/Windows/Fonts/msgothic.ttc"

    image = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)

    left, top = padding, padding
    right, bottom = size - padding, size - padding
    center_x, center_y = size // 2, size // 2

    line_width = 8
    draw.ellipse([left, top, right, bottom], outline=color_vermilion, width=line_width)

    line_y1 = size * 0.32
    line_y2 = size * 0.68
    radius = (size - 2 * padding) / 2

    def get_x_for_y(y):
        y_offset = abs(y - center_y)
        if y_offset >= radius:
            return center_x
        return (radius**2 - y_offset**2)**0.5

    x_offset1 = get_x_for_y(line_y1)
    draw.line([center_x - x_offset1, line_y1, center_x + x_offset1, line_y1], fill=color_vermilion, width=line_width)

    x_offset2 = get_x_for_y(line_y2)
    draw.line([center_x - x_offset2, line_y2, center_x + x_offset2, line_y2], fill=color_vermilion, width=line_width)

    now = datetime.now()
    date_str = now.strftime("%Y.%#m.%#d") if os.name == 'nt' else now.strftime("%Y.%-m.%-d")

    try:
        font_large = ImageFont.truetype(font_path, 80)
        font_date = ImageFont.truetype(font_path, 70)
    except Exception:
        font_large = ImageFont.load_default()
        font_date = ImageFont.load_default()

    draw.text((center_x, line_y1 * 0.6), name_top, fill=color_vermilion, font=font_large, anchor="mm")
    draw.text((center_x, center_y), date_str, fill=color_vermilion, font=font_date, anchor="mm")
    draw.text((center_x, line_y2 + (size - padding - line_y2) * 0.45), name_bottom, fill=color_vermilion, font=font_large, anchor="mm")

    image.save(output_path)
    return output_path