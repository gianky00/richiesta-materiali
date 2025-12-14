import os
from PIL import Image, ImageDraw, ImageFont

def create_modern_icon(text, color_bg, color_text, filename):
    """
    Creates a modern flat icon with text and saves it as an ICO file
    containing multiple sizes.
    """
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = []

    # Path to font
    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

    for size in sizes:
        width, height = size

        # Create image with transparent background (or rounded corners)
        # For simplicity in ICO, we'll use a solid background with a slight rounded effect mask if possible,
        # but standard PIL drawing is easier with full fill for clarity.
        # Let's do a nice rounded square.

        img = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Rounded rectangle parameters
        radius = int(width * 0.2)
        rect_coords = [0, 0, width, height]

        # Draw rounded rectangle
        draw.rounded_rectangle(rect_coords, radius=radius, fill=color_bg)

        # Draw Text
        # Load font dynamically based on size
        font_size = int(height * 0.4)
        try:
            font = ImageFont.truetype(font_path, font_size)
        except IOError:
            font = ImageFont.load_default()

        # Calculate text position to center it
        # Using textbbox for newer Pillow versions
        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
        text_w = right - left
        text_h = bottom - top

        text_x = (width - text_w) / 2
        text_y = (height - text_h) / 2 - (height * 0.05) # Slight adjust up

        draw.text((text_x, text_y), text, fill=color_text, font=font)

        # Add a subtle gloss/highlight (top half lighter)
        # This gives a "modern app" feel
        overlay = Image.new('RGBA', size, (255, 255, 255, 0))
        draw_overlay = ImageDraw.Draw(overlay)
        draw_overlay.rounded_rectangle([0, 0, width, height//2], radius=radius, fill=(255, 255, 255, 30))

        # Compose
        img = Image.alpha_composite(img, overlay)

        images.append(img)

    # Save as ICO
    # The first image is the primary one, save_all=True saves the rest
    base_img = images[0]
    base_img.save(filename, format='ICO', sizes=sizes, append_images=images[1:])
    print(f"Generated {filename}")

def main():
    assets_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "assets")
    if not os.path.exists(assets_dir):
        os.makedirs(assets_dir)

    # App Icon: Intelleo Blue (#007bff) with White Text "RDA"
    create_modern_icon(
        text="RDA",
        color_bg=(0, 123, 255, 255), # #007bff
        color_text=(255, 255, 255, 255),
        filename=os.path.join(assets_dir, "app.ico")
    )

    # Setup Icon: A distinct color, maybe Green or Darker Blue, text "SETUP"
    # Or "INST" for Install
    create_modern_icon(
        text="SET",
        color_bg=(40, 167, 69, 255), # #28a745 (Bootstrap Green)
        color_text=(255, 255, 255, 255),
        filename=os.path.join(assets_dir, "setup.ico")
    )

if __name__ == "__main__":
    main()
