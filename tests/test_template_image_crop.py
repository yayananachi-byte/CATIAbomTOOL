"""Verify template-only CATIA screenshot cropping and square placement."""

import importlib.util
import sys
import unittest
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage


SCRIPT = Path(__file__).resolve().parents[1] / "模板转化" / "bom_to_control_table.py"
spec = importlib.util.spec_from_file_location("bom_to_control_table", SCRIPT)
converter = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = converter
spec.loader.exec_module(converter)


def screenshot(positions):
    picture = Image.new("RGB", (1200, 700), (49, 49, 97))
    draw = ImageDraw.Draw(picture)
    for box in positions:
        draw.rectangle(box, fill=(0, 190, 200))
    draw.ellipse((1120, 10, 1190, 90), fill="white")
    draw.line((1150, 660, 1190, 699), fill="white", width=4)
    result = BytesIO()
    picture.save(result, "PNG")
    return result.getvalue()


class TemplateImageCropTests(unittest.TestCase):
    def test_preserves_separated_parts_and_ignores_corner_overlays(self):
        image = screenshot([(250, 200, 390, 430), (680, 260, 850, 480)])
        with Image.open(converter.crop_to_square(image)) as cropped:
            self.assertEqual(cropped.width, cropped.height)
            self.assertLess(cropped.width, 800)
            self.assertEqual(cropped.getpixel((cropped.width // 2, cropped.height // 2))[0], 49)
            colors = list(cropped.get_flattened_data())
            self.assertGreater(sum(1 for p in colors if p[1] == 190), 10000)
            self.assertEqual(sum(1 for p in colors if p == (255, 255, 255)), 0)

    def test_legacy_product_tree_is_excluded(self):
        picture = Image.new("RGB", (1200, 700), (49, 49, 97))
        draw = ImageDraw.Draw(picture)
        for y in range(10, 590, 35):
            draw.rectangle((0, y, 330, y + 12), fill="white")
        draw.rectangle((450, 150, 850, 550), fill=(0, 190, 200))
        data = BytesIO()
        picture.save(data, "PNG")
        with Image.open(converter.crop_to_square(data.getvalue())) as cropped:
            self.assertLess(cropped.width, 650)
            self.assertEqual(sum(1 for color in cropped.get_flattened_data()
                                 if color == (255, 255, 255)), 0)

    def test_unsafe_background_keeps_full_image_in_square(self):
        picture = Image.new("RGB", (600, 300), "white")
        draw = ImageDraw.Draw(picture)
        for x, y, color in ((24, 150, "red"), (576, 150, "green"),
                            (300, 12, "blue"), (300, 288, "yellow")):
            draw.rectangle((x - 10, y - 10, x + 10, y + 10), fill=color)
        data = BytesIO()
        picture.save(data, "PNG")
        with Image.open(converter.crop_to_square(data.getvalue())) as cropped:
            self.assertEqual(cropped.width, cropped.height)
            self.assertGreaterEqual(cropped.width, 600)

    def test_template_image_is_8_57_cm_square_and_has_room(self):
        sheet = Workbook().active
        source = ExcelImage(BytesIO(screenshot([(300, 150, 700, 500)])))
        converter.copy_thumbnail(sheet, source, "AN10")
        inserted = sheet._images[0]
        self.assertEqual((inserted.width, inserted.height),
                         (converter.TEMPLATE_IMAGE_PX, converter.TEMPLATE_IMAGE_PX))
        self.assertEqual(sheet.column_dimensions["AN"].width, 47)
        self.assertTrue(all(sheet.row_dimensions[row].height == 41 for row in range(10, 16)))
        with Image.open(BytesIO(inserted._data())) as cropped:
            self.assertEqual(cropped.width, cropped.height)


if __name__ == "__main__":
    unittest.main()
