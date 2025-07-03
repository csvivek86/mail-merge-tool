#!/usr/bin/env python3
"""
Icon Generator for NSNA Mail Merge Tool
Creates app icons for macOS (.icns) and Windows (.ico)
"""

import os
import sys
from PIL import Image, ImageDraw, ImageFont

def create_base_icon(size=1024, bg_color="#336699", fg_color="#FFFFFF"):
    """Create a base icon with the specified size and colors"""
    # Create a square canvas
    icon = Image.new('RGBA', (size, size), bg_color)
    draw = ImageDraw.Draw(icon)
    
    # Add text
    try:
        # Try to use a nice font if available
        font = ImageFont.truetype("Arial Bold.ttf", size // 4)
    except IOError:
        # Fall back to default
        font = ImageFont.load_default()
    
    # Add letter "N" for NSNA
    text = "N"
    text_width, text_height = draw.textsize(text, font=font)
    position = ((size - text_width) // 2, (size - text_height) // 2)
    draw.text(position, text, fill=fg_color, font=font)
    
    # Add mail merge style element
    padding = size // 10
    line_width = size // 40
    
    # Horizontal lines representing a document
    for i in range(3):
        y_pos = size // 3 + (i * size // 10)
        draw.line([(padding * 2, y_pos), (size - padding * 2, y_pos)], 
                  fill=fg_color, width=line_width)
    
    # Arrow representing merge action
    arrow_points = [
        (size - padding * 2, size // 2),  # Tip
        (size - padding * 3, size // 2 - size // 10),  # Top
        (size - padding * 3, size // 2 + size // 10)   # Bottom
    ]
    draw.polygon(arrow_points, fill=fg_color)
    
    return icon

def create_macos_icns(base_icon, output_path="app_icon.icns"):
    """Create macOS .icns file from the base icon"""
    # MacOS requires multiple sizes
    sizes = [16, 32, 64, 128, 256, 512, 1024]
    icons = {}
    
    for size in sizes:
        icons[size] = base_icon.resize((size, size), Image.LANCZOS)
    
    # We need temporary storage for the icons
    temp_dir = "icon.iconset"
    os.makedirs(temp_dir, exist_ok=True)
    
    # Save each size with the naming convention macOS expects
    for size in sizes:
        icons[size].save(f"{temp_dir}/icon_{size}x{size}.png")
        # Also save the @2x versions
        if size * 2 in sizes:
            icons[size * 2].save(f"{temp_dir}/icon_{size}x{size}@2x.png")
    
    # Use iconutil to create the .icns file (macOS only)
    if sys.platform == "darwin":
        os.system(f"iconutil -c icns {temp_dir} -o {output_path}")
        print(f"Created macOS icon: {output_path}")
    else:
        print("Warning: macOS icon creation requires macOS.")
    
    # Clean up
    for size in sizes:
        os.remove(f"{temp_dir}/icon_{size}x{size}.png")
        if size * 2 in sizes:
            try:
                os.remove(f"{temp_dir}/icon_{size}x{size}@2x.png")
            except:
                pass
    os.rmdir(temp_dir)

def create_windows_ico(base_icon, output_path="app_icon.ico"):
    """Create Windows .ico file from the base icon"""
    # Windows icon sizes
    sizes = [16, 24, 32, 48, 64, 128, 256]
    icons = []
    
    for size in sizes:
        icons.append(base_icon.resize((size, size), Image.LANCZOS))
    
    # Save as .ico file
    icons[0].save(
        output_path,
        format="ICO",
        sizes=[(size, size) for size in sizes],
        append_images=icons[1:]
    )
    print(f"Created Windows icon: {output_path}")

if __name__ == "__main__":
    print("Generating icons for NSNA Mail Merge Tool...")
    
    # Create base icon
    base_icon = create_base_icon(size=1024, 
                               bg_color="#336699",  # Blue background
                               fg_color="#FFFFFF")  # White foreground
    
    # Save as PNG for reference
    base_icon.save("app_icon.png")
    print("Created base icon: app_icon.png")
    
    # Create platform-specific icons
    create_macos_icns(base_icon)
    create_windows_ico(base_icon)
    
    print("Icon generation complete.")
