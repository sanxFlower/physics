from PIL import Image, ImageOps
import os
base = os.path.dirname(__file__)
src = os.path.join(base, '图片1.2')
out = os.path.join(base, '课件素材')
os.makedirs(out, exist_ok=True)

def crop(name, box, outname, pad=0):
    im = Image.open(os.path.join(src, name)).convert('RGB')
    x1,y1,x2,y2 = box
    x1=max(0,x1-pad); y1=max(0,y1-pad); x2=min(im.width,x2+pad); y2=min(im.height,y2+pad)
    im.crop((x1,y1,x2,y2)).save(os.path.join(out,outname), quality=95)

# 课件截图中的教材原图：路线图、直线位移示意、打点计时器实物
crop('2.jpg', (105, 1050, 410, 1750), 'route_map.jpg', 12)
crop('3.jpg', (410, 510, 950, 1050), 'displacement_diagram.jpg', 12)
crop('4.jpg', (75, 700, 470, 1380), 'timer_devices.jpg', 10)
crop('4.jpg', (570, 1390, 1110, 1770), 'timer_operation.jpg', 10)
print(out)
