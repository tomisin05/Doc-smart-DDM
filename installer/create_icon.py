from PIL import Image, ImageDraw

# Create a simple icon
img = Image.new('RGBA', (64, 64), (70, 130, 180, 255))  # Steel blue background
draw = ImageDraw.Draw(img)

# Draw a document icon
draw.rectangle([10, 8, 50, 56], fill='white', outline='black', width=2)
draw.rectangle([15, 15, 45, 18], fill='black')
draw.rectangle([15, 22, 40, 25], fill='black')
draw.rectangle([15, 29, 45, 32], fill='black')
draw.rectangle([15, 36, 35, 39], fill='black')

# Save as ICO
img.save('icon.ico', format='ICO', sizes=[(64, 64), (32, 32), (16, 16)])
print("Icon created: icon.ico")