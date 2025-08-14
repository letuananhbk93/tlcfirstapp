from rembg import remove
from PIL import Image
input_path = 'Images/ICON.JPG'
output_path = 'Images/ICON.png'
inp = Image.open(input_path)
output = remove(inp)
# Tạo nền trắng
#white_bg = Image.new("RGB", output.size, (255, 255, 255))  # Màu trắng

# Dán ảnh đã xóa nền (nền trong suốt) lên ảnh nền trắng
#white_bg.paste(output, mask=output.split()[3])  # Dùng alpha channel làm mask

# Lưu ảnh đầu ra
#white_bg.save(output_path)
output.save(output_path)