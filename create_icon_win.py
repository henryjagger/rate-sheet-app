"""Creates icon.ico for the Windows build (run by GitHub Actions before PyInstaller)."""
from PIL import Image, ImageDraw, ImageFont


def make_frame(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 255))
    draw = ImageDraw.Draw(img)
    font_size = int(size * 0.60)
    font = None
    for path in [
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]:
        try:
            font = ImageFont.truetype(path, font_size)
            break
        except OSError:
            continue
    if font is None:
        font = ImageFont.load_default()

    text = "%"
    bbox = draw.textbbox((0, 0), text, font=font)
    x = (size - (bbox[2] - bbox[0])) // 2 - bbox[0]
    y = (size - (bbox[3] - bbox[1])) // 2 - bbox[1]
    draw.text((x, y), text, fill=(255, 255, 255, 255), font=font)
    return img


sizes = [256, 128, 64, 48, 32, 16]
frames = [make_frame(s) for s in sizes]
frames[0].save(
    "icon.ico",
    format="ICO",
    append_images=frames[1:],
    sizes=[(s, s) for s in sizes],
)
print("icon.ico created")
