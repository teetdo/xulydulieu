import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns 

# --- Phần 1: Đọc dữ liệu từ file Excel ---
print("--- 1. Đọc dữ liệu ---")
excel_file_path = 'social_media_data (1).xlsx'

try:
    df_facebook = pd.read_excel(excel_file_path, sheet_name='Facebook')
    df_tiktok = pd.read_excel(excel_file_path, sheet_name='TikTok')
    df_instagram = pd.read_excel(excel_file_path, sheet_name='Instagram')
    print(f"Đọc thành công dữ liệu từ file '{excel_file_path}'")
except FileNotFoundError:
    print(f"LỖI: Không tìm thấy file '{excel_file_path}'. Hãy chắc chắn file này tồn tại.")
    exit() 
except Exception as e: 
    print(f"LỖI khi đọc file Excel: {e}")
    exit() 

print("\nDataFrame Facebook (5 dòng đầu):")
print(df_facebook.head())
print(f"Facebook có {len(df_facebook)} bài đăng.")

print("\nDataFrame TikTok (5 dòng đầu):")
print(df_tiktok.head())
print(f"TikTok có {len(df_tiktok)} bài đăng.")

print("\nDataFrame Instagram (5 dòng đầu):")
print(df_instagram.head())
print(f"Instagram có {len(df_instagram)} bài đăng.")


# --- Phần 2: Chuẩn bị và Ghép nối dữ liệu ---
print("\n--- 2. Chuẩn bị và Ghép nối dữ liệu ---")

df_facebook['Platform'] = 'Facebook'
df_tiktok['Platform'] = 'TikTok'
df_instagram['Platform'] = 'Instagram'

df_tong_hop = pd.concat([df_facebook, df_tiktok, df_instagram], ignore_index=True)

cols_tuong_tac = ['Likes', 'Shares', 'Comments']
for col in cols_tuong_tac:
    if col in df_tong_hop.columns:
        df_tong_hop[col] = df_tong_hop[col].fillna(0)
    else:
        print(f"CẢNH BÁO: Cột '{col}' không có trong dữ liệu! Các tính toán liên quan có thể không chính xác.")
        df_tong_hop[col] = 0 

df_tong_hop['Tổng tương tác'] = df_tong_hop['Likes'] + df_tong_hop['Shares'] + df_tong_hop['Comments']

df_tong_hop['Tỉ lệ chia sẻ'] = 0.0 
mask_tong_tuong_tac_khac_0 = df_tong_hop['Tổng tương tác'] != 0
df_tong_hop.loc[mask_tong_tuong_tac_khac_0, 'Tỉ lệ chia sẻ'] = \
    (df_tong_hop['Shares'][mask_tong_tuong_tac_khac_0] / df_tong_hop['Tổng tương tác'][mask_tong_tuong_tac_khac_0]) * 100

print("\nDataFrame Tổng hợp (df_tong_hop) sau khi tính toán (5 dòng đầu):")
print(df_tong_hop.head())
print(f"Tổng cộng có {len(df_tong_hop)} bài đăng từ tất cả các nền tảng.")


# --- Phần 3: Thống kê dữ liệu ---
print("\n--- 3. Thống kê dữ liệu ---")

posts_per_platform = df_tong_hop['Platform'].value_counts().reset_index()
posts_per_platform.columns = ['Platform', 'Tổng số bài đăng'] 
print("\nTổng số bài đăng theo từng nền tảng:")
print(posts_per_platform)

avg_interactions_platform = df_tong_hop.groupby('Platform')[['Likes', 'Shares', 'Comments']].mean().reset_index()
print("\nTương tác trung bình cho từng nền tảng:")
print(avg_interactions_platform)

if 'Chủ đề' in df_tong_hop.columns:
    total_interactions_theme_platform = df_tong_hop.groupby(['Platform', 'Chủ đề'])['Tổng tương tác'].sum().reset_index()
    
    total_interactions_theme_platform_pivot = total_interactions_theme_platform.pivot_table(
        index='Chủ đề', columns='Platform', values='Tổng tương tác', fill_value=0 
    )
    print("\nTổng tương tác theo chủ đề và nền tảng:")
    print(total_interactions_theme_platform_pivot)
else:
    print("\nLƯU Ý: Không tìm thấy cột 'Chủ đề' để thống kê theo chủ đề.")
    total_interactions_theme_platform_pivot = pd.DataFrame() 


# --- Phần 4: Lọc dữ liệu ---
print("\n--- 4. Lọc dữ liệu ---")

df_high_interaction = df_tong_hop[df_tong_hop['Tổng tương tác'] > 1000].copy() 
print(f"\nTìm thấy {len(df_high_interaction)} bài đăng có Tổng tương tác > 1000:")
if not df_high_interaction.empty:
    print(df_high_interaction[['Platform', 'Tổng tương tác', 'Likes', 'Shares', 'Comments']].head())

df_high_share_rate = df_tong_hop[df_tong_hop['Tỉ lệ chia sẻ'] > 20].copy()
print(f"\nTìm thấy {len(df_high_share_rate)} bài đăng có Tỉ lệ chia sẻ > 20%:")
if not df_high_share_rate.empty:
    print(df_high_share_rate[['Platform', 'Tỉ lệ chia sẻ', 'Shares', 'Tổng tương tác']].head())

df_loc_all = pd.DataFrame() 
if not df_high_interaction.empty:
    df_high_interaction['Lý do lọc'] = 'Tổng tương tác > 1000'
    df_loc_all = pd.concat([df_loc_all, df_high_interaction])

if not df_high_share_rate.empty:
    df_high_share_rate['Lý do lọc'] = 'Tỉ lệ chia sẻ > 20%'
    df_loc_all = pd.concat([df_loc_all, df_high_share_rate])

if not df_loc_all.empty:
    unique_cols = ['Platform', 'Likes', 'Shares', 'Comments', 'Lý do lọc']
    if 'Post ID' in df_loc_all.columns:
        unique_cols.insert(0, 'Post ID')
    
    valid_unique_cols = [col for col in unique_cols if col in df_loc_all.columns]
    if valid_unique_cols:
         df_loc_all = df_loc_all.drop_duplicates(subset=valid_unique_cols).reset_index(drop=True)


# --- Phần 5: Vẽ biểu đồ ---
print("\n--- 5. Vẽ biểu đồ ---")

if not df_tong_hop.empty:
    # 1. Biểu đồ scatter: Mối quan hệ giữa 'Tổng tương tác' và 'Tỉ lệ chia sẻ'
    plt.figure(figsize=(10, 6)) 
    sns.scatterplot(data=df_tong_hop, x='Tổng tương tác', y='Tỉ lệ chia sẻ', hue='Platform')
    plt.title('Tổng tương tác vs. Tỉ lệ chia sẻ (theo Nền tảng)') 
    plt.xlabel('Tổng số tương tác') 
    plt.ylabel('Tỉ lệ chia sẻ (%)') 
    plt.grid(True) 
    plt.legend(title='Nền tảng') 
    plt.show() 

    # 2. Biểu đồ cột: So sánh tổng tương tác giữa các nền tảng
    total_interactions_by_platform_sum = df_tong_hop.groupby('Platform')['Tổng tương tác'].sum().reset_index()
    if not total_interactions_by_platform_sum.empty:
        plt.figure(figsize=(7, 5))
        sns.barplot(data=total_interactions_by_platform_sum, x='Platform', y='Tổng tương tác', palette=['skyblue', 'lightcoral', 'lightgreen'])
        plt.title('Tổng tương tác của mỗi Nền tảng')
        plt.xlabel('Nền tảng')
        plt.ylabel('Tổng số lượt tương tác')
        plt.show()
else:
    print("Không có dữ liệu để vẽ biểu đồ.")

print("\n--- 6. Xuất dữ liệu ra Excel ---")
output_excel_path = 'social_media_analysis.xlsx' 

try:
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        # Sheet 1: Dữ liệu tổng hợp
        df_tong_hop.to_excel(writer, sheet_name='DuLieuTongHop', index=False) 

        # Sheet 2: Kết quả thống kê
        start_row = 0 

        pd.DataFrame(["Tổng số bài đăng theo từng nền tảng:"]).to_excel(writer, sheet_name='KetQuaThongKe', startrow=start_row, index=False, header=False)
        start_row += 2 
        posts_per_platform.to_excel(writer, sheet_name='KetQuaThongKe', startrow=start_row, index=False)
        start_row += len(posts_per_platform) + 2 

        pd.DataFrame(["Tương tác trung bình cho từng nền tảng:"]).to_excel(writer, sheet_name='KetQuaThongKe', startrow=start_row, index=False, header=False)
        start_row += 2
        avg_interactions_platform.to_excel(writer, sheet_name='KetQuaThongKe', startrow=start_row, index=False)
        start_row += len(avg_interactions_platform) + 2

        if not total_interactions_theme_platform_pivot.empty:
            pd.DataFrame(["Tổng tương tác theo chủ đề và nền tảng:"]).to_excel(writer, sheet_name='KetQuaThongKe', startrow=start_row, index=False, header=False)
            start_row += 2
            total_interactions_theme_platform_pivot.reset_index().to_excel(writer, sheet_name='KetQuaThongKe', startrow=start_row, index=False)

        # Sheet 3: Dữ liệu đã lọc
        if not df_loc_all.empty:
            df_loc_all.to_excel(writer, sheet_name='DuLieuDaLoc', index=False)
        else:
            pd.DataFrame(["Không có dữ liệu nào được lọc."]).to_excel(writer, sheet_name='DuLieuDaLoc', index=False, header=False)

    print(f"\nĐã xuất kết quả phân tích ra file: '{output_excel_path}'")
    print("File kết quả bao gồm các sheet: 'DuLieuTongHop', 'KetQuaThongKe', 'DuLieuDaLoc'")

except Exception as e:
    print(f"LỖI khi xuất file Excel '{output_excel_path}': {e}")

print("\n--- Hoàn thành! ---")