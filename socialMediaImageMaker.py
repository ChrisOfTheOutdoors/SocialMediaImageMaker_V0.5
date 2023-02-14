import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import tkinter as tk
from tkinter import filedialog
import io

# Initialize the GUI window
root = tk.Tk()
root.withdraw()

# Ask user to select the input Excel file
file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx;*.xls")])

# Ask user to select the output directory for the generated images
output_dir = filedialog.askdirectory(title="Select output directory")

# Read the Excel file into a Pandas DataFrame
df = pd.read_excel(file_path, usecols=["Title", "Content", "Image"])

# Create the output directory if it doesn't exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Define image size and font settings
image_width = 500
image_height = 500
font_file = 'arial.ttf'
title_font_size = 48
content_font_size = 36
text_color = (255, 255, 255)

# Loop through each row in the DataFrame and generate an image for it
for i, row in df.iterrows():
    # Get the image path and open it with PIL
    image_path = row['Image']
    image = Image.open(image_path)

    # Resize the image to fit the desired dimensions
    image = image.resize((image_width, image_height))

    # Create a new image with a white background and the resized image pasted on it
    background_image = Image.new('RGB', (image_width, image_height), 'white')
    background_image.paste(image, (0, 0))

    # Initialize a drawing context on the image
    draw = ImageDraw.Draw(background_image)

    # Add the title text to the image
    title = row['Title']
    title_font = ImageFont.truetype(font_file, title_font_size)
    title_text_size = draw.textbbox((0, 0, image_width, image_height), title, font=title_font)
    title_text_position = ((image_width - title_text_size[2]) / 2, 50)
    draw.text(title_text_position, title, font=title_font, fill=text_color)

    # Add the content text to the image
    content = row['Content']
    content_font = ImageFont.truetype(font_file, content_font_size)
    content_text_size = draw.textbbox((0, 0, image_width, image_height), content, font=content_font)
    content_text_position = ((image_width - content_text_size[2]) / 2, 200)
    draw.text(content_text_position, content, font=content_font, fill=text_color)

    # Save the image to the output directory
    image_filename = f"image_{i+1}.png"
    image_path = os.path.join(output_dir, image_filename)
    background_image.save(image_path)
