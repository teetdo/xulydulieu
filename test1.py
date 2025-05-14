import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

# --- a. Chuẩn bị dữ liệu ---
print("--- a. Chuẩn bị dữ liệu ---")

# Hàm tạo dữ liệu mẫu
def generate_data(platform_name, num_records=60):
    data = []
    start_date = datetime(2023, 1, 1)
    themes = ["Giải trí", "Giáo dục", "Thể thao", "Công nghệ", "Đời sống", "Du lịch", "Ẩm thực"]
    for i in range(1, num_records + 1):
        post_id = f"{platform_name[0]}{i:03d}"
        content = f"Nội dung bài đăng {platform_name} số {i} về {random.choice(themes).lower()}."
        date = start_date + timedelta(days=random.randint(0, 365), hours=random.randint(0,23), minutes=random.randint(0,59))
        likes = random.randint(50, 5000)
        shares = random.randint(0, likes // 2) # Shares thường ít hơn likes
        comments = random.randint(0, likes // 5) # Comments thường ít hơn shares
        theme = random.choice(themes)
        
        if platform_name == "TikTok":
            views = random.randint(likes * 5, likes * 50) # TikTok thường có views cao
            data.append([post_id, content, date.strftime("%Y-%m-%d %H:%M:%S"), views, likes, shares, comments, theme])
        else:
            data.append([post_id, content, date.strftime("%Y-%m-%d %H:%M:%S"), likes, shares, comments, theme])
    
    if platform_name == "TikTok":
        columns = ["Post ID", "Nội dung", "Ngày đăng", "Lượt xem", "Likes", "Shares", "Comments", "Chủ đề"]
    else:
        columns = ["Post ID", "Nội dung", "Ngày đăng", "Likes", "Shares", "Comments", "Chủ đề"]
        
    return pd.DataFrame(data, columns=columns)

# Tạo dữ liệu cho từng sheet
df_facebook_raw = generate_data("Facebook", 60)
df_tiktok_raw = generate_data("TikTok", 70) # TikTok có thêm cột Lượt xem
df_instagram_raw = generate_data("Instagram", 55)

# Xuất ra file Excel mẫu
excel_file_path = 'social_media_sample.xlsx'
with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
    df_facebook_raw.to_excel(writer, sheet_name='Facebook', index=False)
    df_tiktok_raw.to_excel(writer, sheet_name='TikTok', index=False)
    df_instagram_raw.to_excel(writer, sheet_name='Instagram', index=False)

print(f"Đã tạo file Excel mẫu '{excel_file_path}' với dữ liệu.")
print(f"Facebook: {len(df_facebook_raw)} bản ghi")
print(f"TikTok: {len(df_tiktok_raw)} bản ghi")
print(f"Instagram: {len(df_instagram_raw)} bản ghi")