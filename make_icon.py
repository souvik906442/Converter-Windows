from PIL import Image, ImageDraw
S = 256
im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
d.rounded_rectangle([0, 0, S - 1, S - 1], radius=52, fill=(21, 101, 192, 255))
d.polygon([(66, 44), (150, 44), (184, 78), (184, 206), (66, 206)], fill=(255, 255, 255, 255))
d.polygon([(150, 44), (150, 78), (184, 78)], fill=(187, 222, 251, 255))
for y, w in ((100, 90), (128, 90), (156, 54)):
    d.rectangle([84, y, 84 + w, y + 9], fill=(21, 101, 192, 255))
d.ellipse([126, 126, 226, 226], fill=(255, 152, 0, 255))
d.line([(148, 176), (204, 176)], fill="white", width=10)
d.line([(184, 154), (206, 176), (184, 198)], fill="white", width=10, joint="curve")
im.save("icon.ico", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
