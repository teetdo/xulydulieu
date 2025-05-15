import pandas as pd
import numpy as np
import seaborn as sns

import matplotlib.pyplot as plt

print("--- Bắt đầu Phần 1: Đọc dữ liệu từ file Excel ---")

ten_file_excel = 'social_media_data (1).xlsx'

try:
    df_facebook = pd.read_excel(ten_file_excel, sheet_name='Facebook')
    df_tiktok = pd.read_excel(ten_file_excel, sheet_name='TikTok')
    df_instagram = pd.read_excel(ten_file_excel, sheet_name='Instagram')
    print(f"Đã đọc thành công dữ liệu từ file: '{ten_file_excel}'")
except FileNotFoundError:
    print(f"LỖI: Không tìm thấy file '{ten_file_excel}'. Bạn kiểm tra lại tên file và vị trí nhé.")
    exit()
except Exception as loi_khac:
    print(f"LỖI không mong muốn khi đọc file Excel: {loi_khac}")
    exit()

print("\n--- Xem thử 5 dòng đầu của mỗi bảng ---")
print("Bảng Facebook:")
print(df_facebook.head())
print(f"Số bài đăng trên Facebook: {len(df_facebook)} bài")

print("\nBảng TikTok:")
print(df_tiktok.head())
print(f"Số bài đăng trên TikTok: {len(df_tiktok)} bài")

print("\nBảng Instagram:")
print(df_instagram.head())
print(f"Số bài đăng trên Instagram: {len(df_instagram)} bài")

print("\n--- Bắt đầu Phần 2: Chuẩn bị và ghép nối dữ liệu ---")

df_facebook['Platform'] = 'Facebook'
df_tiktok['Platform'] = 'TikTok'
df_instagram['Platform'] = 'Instagram'

df_tong_hop = pd.concat([df_facebook, df_tiktok, df_instagram], ignore_index=True)

cac_cot_tuong_tac = ['Likes', 'Shares', 'Comments']
for ten_cot in cac_cot_tuong_tac:
    if ten_cot in df_tong_hop.columns:
        df_tong_hop[ten_cot] = df_tong_hop[ten_cot].fillna(0)
    else:
        print(f"CẢNH BÁO: Không tìm thấy cột '{ten_cot}'. Sẽ tạo cột này với giá trị 0.")
        df_tong_hop[ten_cot] = 0

df_tong_hop['Tổng tương tác'] = df_tong_hop['Likes'] + df_tong_hop['Shares'] + df_tong_hop['Comments']
df_tong_hop['Tỉ lệ chia sẻ'] = 0.0
nhung_dong_co_tuong_tac = (df_tong_hop['Tổng tương tác'] != 0)
df_tong_hop.loc[nhung_dong_co_tuong_tac, 'Tỉ lệ chia sẻ'] = \
    (df_tong_hop['Shares'][nhung_dong_co_tuong_tac] / df_tong_hop['Tổng tương tác'][nhung_dong_co_tuong_tac]) * 100

print("\n--- Xem thử bảng tổng hợp sau khi tính toán (5 dòng đầu) ---")
print(df_tong_hop.head())
print(f"Tổng cộng có {len(df_tong_hop)} bài đăng từ tất cả các nền tảng.")

print("\n--- Bắt đầu Phần 3: Thống kê dữ liệu ---")

thong_ke_bai_dang = df_tong_hop['Platform'].value_counts().reset_index()
thong_ke_bai_dang.columns = ['Platform', 'Tổng số bài đăng']
print("\nTổng số bài đăng của mỗi nền tảng:")
print(thong_ke_bai_dang)

thong_ke_trung_binh = df_tong_hop.groupby('Platform')[['Likes', 'Shares', 'Comments']].mean().reset_index()
print("\nTương tác trung bình của mỗi nền tảng:")
print(thong_ke_trung_binh)

if 'Chủ đề' in df_tong_hop.columns:
    thong_ke_chu_de = df_tong_hop.groupby(['Platform', 'Chủ đề'])['Tổng tương tác'].sum().reset_index()
    thong_ke_chu_de_pivot = thong_ke_chu_de.pivot_table(
        index='Chủ đề',
        columns='Platform',
        values='Tổng tương tác',
        fill_value=0
    )
    print("\nTổng tương tác theo từng chủ đề và nền tảng:")
    print(thong_ke_chu_de_pivot)
else:
    print("\nLƯU Ý: Không tìm thấy cột 'Chủ đề' trong dữ liệu, nên bỏ qua thống kê này.")
    thong_ke_chu_de_pivot = pd.DataFrame()

print("\n--- Bắt đầu Phần 4: Lọc dữ liệu ---")

bai_dang_tuong_tac_cao = df_tong_hop[df_tong_hop['Tổng tương tác'] > 1000].copy()
print(f"\nTìm thấy {len(bai_dang_tuong_tac_cao)} bài đăng có Tổng tương tác > 1000.")
if not bai_dang_tuong_tac_cao.empty:
    print("5 bài đầu tiên có tương tác cao:")
    print(bai_dang_tuong_tac_cao[['Platform', 'Tổng tương tác', 'Likes']].head())

bai_dang_share_cao = df_tong_hop[df_tong_hop['Tỉ lệ chia sẻ'] > 20].copy()
print(f"\nTìm thấy {len(bai_dang_share_cao)} bài đăng có Tỉ lệ chia sẻ > 20%.")
if not bai_dang_share_cao.empty:
    print("5 bài đầu tiên có tỉ lệ chia sẻ cao:")
    print(bai_dang_share_cao[['Platform', 'Tỉ lệ chia sẻ', 'Shares']].head())

danh_sach_bang_da_loc = []

if not bai_dang_tuong_tac_cao.empty:
    bai_dang_tuong_tac_cao['Lý do lọc'] = 'Tổng tương tác > 1000'
    danh_sach_bang_da_loc.append(bai_dang_tuong_tac_cao)

if not bai_dang_share_cao.empty:
    bai_dang_share_cao['Lý do lọc'] = 'Tỉ lệ chia sẻ > 20%'
    danh_sach_bang_da_loc.append(bai_dang_share_cao)

if danh_sach_bang_da_loc:
    df_tat_ca_da_loc = pd.concat(danh_sach_bang_da_loc, ignore_index=True)
    cot_kiem_tra_lap = ['Platform', 'Likes', 'Shares', 'Comments', 'Lý do lọc']
    if 'Post ID' in df_tat_ca_da_loc.columns:
        cot_kiem_tra_lap.insert(0, 'Post ID')
    cot_kiem_tra_lap_hop_le = [cot for cot in cot_kiem_tra_lap if cot in df_tat_ca_da_loc.columns]
    if cot_kiem_tra_lap_hop_le:
        df_tat_ca_da_loc = df_tat_ca_da_loc.drop_duplicates(subset=cot_kiem_tra_lap_hop_le).reset_index(drop=True)
else:
    df_tat_ca_da_loc = pd.DataFrame()

print("\n--- Bắt đầu Phần 5: Vẽ biểu đồ ---")

if not df_tong_hop.empty:
    plt.figure(figsize=(10, 6)) 
    sns.scatterplot(data=df_tong_hop, x='Tổng tương tác', y='Tỉ lệ chia sẻ', hue='Platform')
    plt.title('Biểu đồ 1: Tổng tương tác và Tỉ lệ chia sẻ') 
    plt.xlabel('Tổng số tương tác của bài đăng')      
    plt.ylabel('Tỉ lệ chia sẻ của bài đăng (%)')    
    plt.grid(True) 
    plt.legend(title='Nền tảng') 

    tong_tuong_tac_theo_nen_tang = df_tong_hop.groupby('Platform')['Tổng tương tác'].sum().reset_index()
    if not tong_tuong_tac_theo_nen_tang.empty:
        plt.figure(figsize=(7, 5)) 
        sns.barplot(data=tong_tuong_tac_theo_nen_tang, x='Platform', y='Tổng tương tác', palette=['orange', 'green', 'purple'])
        plt.title('Biểu đồ 2: Tổng tương tác của mỗi Nền tảng')
        plt.xlabel('Nền tảng')
        plt.ylabel('Tổng số lượt tương tác')
    
    plt.show() 

else:
    print("Không có dữ liệu trong bảng tổng hợp để vẽ biểu đồ.")

print("\n--- Bắt đầu Phần 6: Xuất dữ liệu ra file Excel ---")
ten_file_ket_qua = 'social_media_analysis1.xlsx'

try:
    with pd.ExcelWriter(ten_file_ket_qua, engine='openpyxl') as writer:
        df_tong_hop.to_excel(writer, sheet_name='DuLieuGhepLai', index=False)
        dong_bat_dau_ghi = 0
        pd.DataFrame(["Thống kê: Tổng số bài đăng theo nền tảng"]).to_excel(writer, sheet_name='CacBangThongKe', startrow=dong_bat_dau_ghi, index=False, header=False)
        dong_bat_dau_ghi += 2
        thong_ke_bai_dang.to_excel(writer, sheet_name='CacBangThongKe', startrow=dong_bat_dau_ghi, index=False)
        dong_bat_dau_ghi += len(thong_ke_bai_dang) + 2
        pd.DataFrame(["Thống kê: Tương tác trung bình theo nền tảng"]).to_excel(writer, sheet_name='CacBangThongKe', startrow=dong_bat_dau_ghi, index=False, header=False)
        dong_bat_dau_ghi += 2
        thong_ke_trung_binh.to_excel(writer, sheet_name='CacBangThongKe', startrow=dong_bat_dau_ghi, index=False)
        dong_bat_dau_ghi += len(thong_ke_trung_binh) + 2
        if not thong_ke_chu_de_pivot.empty:
            pd.DataFrame(["Thống kê: Tổng tương tác theo chủ đề và nền tảng"]).to_excel(writer, sheet_name='CacBangThongKe', startrow=dong_bat_dau_ghi, index=False, header=False)
            dong_bat_dau_ghi += 2
            thong_ke_chu_de_pivot.reset_index().to_excel(writer, sheet_name='CacBangThongKe', startrow=dong_bat_dau_ghi, index=False)
        if not df_tat_ca_da_loc.empty:
            df_tat_ca_da_loc.to_excel(writer, sheet_name='DuLieuDaLoc', index=False)
        else:
            pd.DataFrame(["Không có bài đăng nào được lọc theo tiêu chí."]).to_excel(writer, sheet_name='DuLieuDaLoc', index=False, header=False)
    print(f"\nĐã xuất tất cả kết quả phân tích ra file Excel: '{ten_file_ket_qua}'")
    print("File này sẽ có các sheet: 'DuLieuGhepLai', 'CacBangThongKe', 'DuLieuDaLoc'")
except Exception as loi_xuat_file:
    print(f"LỖI: Không thể xuất dữ liệu ra file Excel. Chi tiết lỗi: {loi_xuat_file}")

print("\n--- Chương trình đã chạy xong! ---")
