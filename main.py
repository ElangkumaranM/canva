from PIL import Image, ImageFont, ImageDraw
import pandas as pd
import os

# Set up font, font color, and certificate dimensions
FONT_FILE = r'C:\Users\Elangkumaran M\Desktop\C\certificate-generator\Roboto-BlackItalic.ttf'  # Path to font file
FONT_COLOR = "#1E49E2"
WIDTH, HEIGHT = 720, 509  # Dimensions of your certificate image

def make_cert(name):
    """Function to generate a certificate with the given name."""
    # Load the certificate template image
    image_source = Image.open(r'C:\Users\Elangkumaran M\Desktop\C\certificate-generator\canva.jpg')  # Path to uploaded image
    # Create a draw object
    draw = ImageDraw.Draw(image_source)
    # Load the font
    font = ImageFont.truetype(FONT_FILE, 30)  # Adjust font size as needed
    
    text = name  # Text to be drawn
    # Calculate text bounding box and dimensions
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]

    # Calculate position to center the text horizontally and align vertically under "THIS IS PRESENTED TO"
    text_x = (WIDTH - text_width) / 3
    # Adjust the Y position to place it under the "This is presented to" line
    text_y = 260  # Adjust the Y value based on the desired position
    
    # Draw the name on the certificate
    draw.text((text_x, text_y), text, fill=FONT_COLOR, font=font)

    # Save the certificate as a PNG file
    output_file = os.path.join(output_dir, f"{name}.png")  # Use os.path.join for better path handling
    image_source.save(output_file, format="PNG")

    return output_file

# Ensure the output directory exists
output_dir = './out/'  # Directory to save the generated certificates
os.makedirs(output_dir, exist_ok=True)

# Read names from an Excel file (assuming the names are in a column named 'Name')
df = pd.read_excel(r'C:\Users\Elangkumaran M\Desktop\C\certificate-generator\data.xlsx')
names = df['name'].tolist()

# Generate and save certificates
for name in names:
    make_cert(name)
