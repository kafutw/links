"""把去背商品圖放到同一個 4:5 透明畫布上，讓每個商品看起來差不多大。
做法：商品不能超出畫布的 92%；在這個上限內，把「商品實際佔的面積」盡量拉到畫布的 TARGET 比例。
用法：python3 normalize_thumb.py 輸入.png 輸出.png
"""
import sys
from PIL import Image

W, H = 400, 500
MAX_FILL = 0.92
TARGET = 0.50

src = Image.open(sys.argv[1]).convert('RGBA')
src = src.crop(src.getchannel('A').getbbox())
alpha = src.getchannel('A').point(lambda v: 255 if v > 24 else 0)
coverage = sum(alpha.histogram()[255:]) / (src.width * src.height)   # 商品在自己外框裡佔多少

fit = min(W * MAX_FILL / src.width, H * MAX_FILL / src.height)       # 不超出畫布的最大倍率
want = (TARGET * W * H / (src.width * src.height * coverage)) ** 0.5  # 讓面積達標的倍率
scale = min(fit, want)
im = src.resize((round(src.width * scale), round(src.height * scale)), Image.LANCZOS)
canvas = Image.new('RGBA', (W, H), (0, 0, 0, 0))
canvas.alpha_composite(im, ((W - im.width) // 2, (H - im.height) // 2))
canvas.save(sys.argv[2], optimize=True)
area = sum(im.getchannel('A').point(lambda v: 255 if v > 24 else 0).histogram()[255:]) / (W * H)
print(f'{sys.argv[2]}: 商品 {im.width}x{im.height}，佔畫布 {area:.0%}')
