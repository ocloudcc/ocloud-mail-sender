from PIL import Image

# 打开你的图片
img = Image.open('app.png')
# 保存为包含多分辨率的 ico 文件
img.save(
    'app.ico',
    format='ICO',
    sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
)
print('已生成 app.ico')