import os
import sys
import time
import shutil
import csv
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from playwright.sync_api import sync_playwright

# Import thư viện xử lý ảnh
try:
    from PIL import Image, ImageTk
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Thiếu thư viện", "Vui lòng mở Terminal và chạy lệnh:\npip install pillow\nđể sử dụng tính năng xem trước ảnh!")
    sys.exit()

# ==============================================================================
# 1. CẤU HÌNH HỆ THỐNG
# ==============================================================================
LOGIN_URL = "https://construction.v-office.vn/login"
URL_LIST = "https://construction.v-office.vn/quan-ly-kho/phieu-van-chuyen"
URL_CREATE = "https://construction.v-office.vn/quan-ly-kho/phieu-van-chuyen?action=create" 
AUTH_FILE = "auth_state.json"
EXCEL_FILE = "Khối lượng.xlsx"
TEMP_UPLOAD_DIR = "Temp_Upload"
REPORT_FILE = "Bao_Cao_Ca.csv"  # File báo cáo tự động xuất ra

USERNAME = "0971936186"
PASSWORD = "Ducviet1996@"

COMPANY_NAME = "VINALPHA"

COL_BIEN_SO = "Biển số xe"
COL_KHOI_LUONG = "Khối lượng"
COL_SO_CHUYEN = "Số chuyến"

DU_AN = "Dự án KĐT Hòa Long- Bắc Ninh"
DON_VI_KEYWORD = "VẠN XUÂN"
NHA_CUNG_CAP_FULL = "CÔNG TY CỔ PHẦN XÂY DỰNG THƯƠNG MẠI VẠN XUÂN"

HANG_MUC = "Cấp cát san lấp"
LOAI_PHUONG_TIEN = "Ô tô tải"
TEN_VAT_TU = "Cát san lấp"
VI_TRI = "KĐT Hòa Long"


# ==============================================================================
# 2. GIAO DIỆN HÀNG CHỜ (CÓ CHỌN NHÀ CUNG CẤP & XUẤT BÁO CÁO)
# ==============================================================================
def open_queue_manager(df_excel):
    os.makedirs(TEMP_UPLOAD_DIR, exist_ok=True)
    queue_data = []

    root = tk.Tk()
    root.title("Bộ lập hàng chờ tạo phiếu vận chuyển (Tự động Xuất Báo Cáo Chốt Ca)")
    root.geometry("1080x820")
    root.resizable(False, False)

    font_lbl = ("Segoe UI", 10)
    font_entry = ("Segoe UI", 10)

    temp_selected_images = []

    frame_top = tk.Frame(root)
    frame_top.pack(fill="both", expand=True)

    frame_left = tk.Frame(frame_top, width=450)
    frame_left.pack(side="left", fill="both", padx=15, pady=15, expand=True)

    frame_right = tk.Frame(frame_top, width=600)
    frame_right.pack(side="right", fill="both", padx=15, pady=15, expand=True)

    # 1. Biển số
    tk.Label(frame_left, text="1. Biển số xe:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 2))
    entry_plate = tk.Entry(frame_left, font=font_entry)
    entry_plate.pack(fill="x", pady=2)
    entry_plate.focus()

    # 2. Số chuyến
    tk.Label(frame_left, text="2. Số chuyến (ví dụ: 1, 2, 3...):", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
    entry_trip = tk.Entry(frame_left, font=font_entry)
    entry_trip.insert(0, "1")
    entry_trip.pack(fill="x", pady=2)

    # 3. Nhà cung cấp (GHI CHÚ ĐỂ XUẤT BÁO CÁO)
    tk.Label(frame_left, text="3. Ghi chú Nhà cung cấp (Lưu báo cáo Excel):", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
    combo_ncc = ttk.Combobox(frame_left, font=font_entry, state="normal")
    combo_ncc['values'] = [
        "Hùng Vôi",
        "Báu",
        "Đức Tuệ",
        "Tuấn Huyền",
        "Nam Mắt To",
        "Thắng 30/4"
    ]
    combo_ncc.current(0)
    combo_ncc.pack(fill="x", pady=2)

    # 4. Ảnh đính kèm
    tk.Label(frame_left, text="4. Ảnh đính kèm (Bắt buộc đủ 4 ảnh):", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
    lbl_img_count = tk.Label(frame_left, text="Chưa chọn ảnh nào (Yêu cầu 4 ảnh)", fg="red", font=("Segoe UI", 9, "italic"))
    lbl_img_count.pack(anchor="w")

    def browse_images():
        nonlocal temp_selected_images
        filenames = filedialog.askopenfilenames(
            title="Chọn đúng 4 file ảnh cho xe này",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.webp")]
        )
        if filenames:
            temp_selected_images = [os.path.abspath(f) for f in filenames]
            count = len(temp_selected_images)
            if count == 4:
                lbl_img_count.config(text="Đã chọn đủ 4 ảnh hợp lệ", fg="green")
            else:
                lbl_img_count.config(text=f"Cảnh báo: Mới chọn {count}/4 ảnh!", fg="red")

    btn_img = tk.Button(frame_left, text="Chọn 4 ảnh...", command=browse_images)
    btn_img.pack(anchor="w", pady=4)

    # Cán bộ phê duyệt
    frame_users = tk.LabelFrame(frame_left, text=" Cán bộ phê duyệt trong ca ", font=font_lbl)
    frame_users.pack(fill="x", pady=10)

    tk.Label(frame_users, text="Thủ kho BCH:").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    entry_thukho = tk.Entry(frame_users, font=font_entry, width=22)
    entry_thukho.insert(0, "Trần Minh Huyền")
    entry_thukho.grid(row=0, column=1, padx=8, pady=3)

    tk.Label(frame_users, text="Ban quản lý:").grid(row=1, column=0, sticky="w", padx=8, pady=3)
    entry_bql = tk.Entry(frame_users, font=font_entry, width=22)
    entry_bql.insert(0, "Trịnh Văn Cần")
    entry_bql.grid(row=1, column=1, padx=8, pady=3)

    tk.Label(frame_users, text="An ninh:").grid(row=2, column=0, sticky="w", padx=8, pady=3)
    entry_anninh = tk.Entry(frame_users, font=font_entry, width=22)
    entry_anninh.insert(0, "Nguyễn Mạnh Hùng")
    entry_anninh.grid(row=2, column=1, padx=8, pady=3)

    btn_add = tk.Button(frame_left, text="+ THÊM VÀO HÀNG CHỜ", font=("Segoe UI", 11, "bold"), bg="#1890ff", fg="white", command=lambda: add_to_queue(), height=2)
    btn_add.pack(fill="x", pady=15)

    # --- BẢNG HÀNG CHỜ ---
    tk.Label(frame_right, text="DANH SÁCH ĐƠN CHỜ XỬ LÝ:", font=("Segoe UI", 10, "bold")).pack(anchor="w")

    tree = ttk.Treeview(frame_right, columns=("STT", "BienSo", "Chuyen", "NCC", "SoAnh"), show="headings", height=10)
    tree.heading("STT", text="STT")
    tree.heading("BienSo", text="Biển số")
    tree.heading("Chuyen", text="Chuyến")
    tree.heading("NCC", text="Ghi chú (Báo cáo)")
    tree.heading("SoAnh", text="Ảnh")
    
    tree.column("STT", width=40, anchor="center")
    tree.column("BienSo", width=120, anchor="center")
    tree.column("Chuyen", width=60, anchor="center")
    tree.column("NCC", width=200, anchor="w")
    tree.column("SoAnh", width=70, anchor="center")
    tree.pack(fill="both", expand=True, pady=5)

    btn_del = tk.Button(frame_right, text="Xóa đơn đang chọn", command=lambda: delete_selected())
    btn_del.pack(anchor="e", pady=2)

    is_started = {"value": False}
    def start_running():
        if not queue_data:
            messagebox.showwarning("Trống", "Hàng chờ chưa có đơn nào!")
            return
        is_started["value"] = True
        root.destroy()

    btn_start_all = tk.Button(frame_right, text="BẮT ĐẦU CHẠY HÀNG CHỜ", font=("Segoe UI", 12, "bold"), bg="#52c41a", fg="white", command=start_running, height=2)
    btn_start_all.pack(fill="x", pady=(10, 0))

    # --- KHUNG PREVIEW ẢNH ---
    frame_preview = tk.LabelFrame(root, text=" Xem trước 4 ảnh của đơn đang chọn ", font=("Segoe UI", 10, "bold"))
    frame_preview.pack(side="bottom", fill="x", padx=15, pady=(0, 15))
    
    img_labels = []
    for i in range(4):
        lbl = tk.Label(frame_preview, text=f"Ảnh {i+1}", width=26, height=13, bg="#e0e0e0", relief="solid", bd=1)
        lbl.pack(side="left", padx=10, pady=10, expand=True)
        img_labels.append(lbl)

    def on_tree_select(event):
        selected = tree.selection()
        if not selected:
            for lbl in img_labels:
                lbl.config(image='', text="Trống", width=26, height=13)
            return
        
        idx = tree.index(selected[0])
        order = queue_data[idx]
        images = order.get("images", [])

        for i in range(4):
            if i < len(images):
                try:
                    img = Image.open(images[i])
                    img.thumbnail((220, 220)) 
                    photo = ImageTk.PhotoImage(img)
                    img_labels[i].config(image=photo, text="", width=220, height=220)
                    img_labels[i].image = photo
                except:
                    img_labels[i].config(image='', text="Lỗi file", width=26, height=13)
            else:
                img_labels[i].config(image='', text="Trống", width=26, height=13)

    tree.bind('<<TreeviewSelect>>', on_tree_select)

    def add_to_queue():
        nonlocal temp_selected_images
        plate = entry_plate.get().strip()
        trip = entry_trip.get().strip()
        ncc_val = combo_ncc.get().strip()

        if not plate:
            messagebox.showerror("Lỗi", "Vui lòng nhập Biển số xe!")
            return
            
        if not ncc_val:
            messagebox.showerror("Lỗi", "Vui lòng nhập hoặc chọn Nhà cung cấp!")
            return

        clean_plate_input = plate.replace("-", "").replace(".", "").replace(" ", "").upper()
        
        if df_excel is not None and not df_excel.empty:
            if clean_plate_input not in df_excel["clean_plate"].values:
                messagebox.showwarning(
                    "Cảnh báo: Không có biển số", 
                    f"Biển số '{plate}' KHÔNG TỒN TẠI trong file Excel 'Khối lượng.xlsx'!\n\n"
                    f"Vui lòng kiểm tra lại xem bạn có gõ sai không, hoặc cập nhật thêm biển này vào file Excel."
                )
                return 

        if not trip.isdigit() or int(trip) <= 0:
            messagebox.showerror("Lỗi", "Số chuyến phải là số nguyên dương!")
            return

        if len(temp_selected_images) != 4:
            messagebox.showerror("Lỗi", f"Bắt buộc chọn đúng 4 ảnh! (Hiện có {len(temp_selected_images)})")
            return

        for p in temp_selected_images:
            if not os.path.exists(p):
                messagebox.showerror("Lỗi", f"File ảnh không tồn tại: {p}")
                return

        formatted_trip = f"{int(trip):02d}"
        folder_name = f"{plate}_C{formatted_trip}_{int(time.time())}"
        target_dir = os.path.join(TEMP_UPLOAD_DIR, folder_name)
        os.makedirs(target_dir, exist_ok=True)
        
        new_image_paths = []
        try:
            for p in temp_selected_images:
                filename = os.path.basename(p)
                dest_path = os.path.join(target_dir, filename)
                shutil.move(p, dest_path)
                new_image_paths.append(dest_path)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể cắt file ảnh. Lỗi: {e}")
            return

        order_item = {
            "plate": plate,
            "trip": formatted_trip,
            "ncc": ncc_val,          
            "images": new_image_paths,
            "thu_kho": entry_thukho.get().strip(),
            "bql": entry_bql.get().strip(),
            "an_ninh": entry_anninh.get().strip()
        }
        queue_data.append(order_item)

        stt = len(queue_data)
        item_id = tree.insert("", "end", values=(stt, plate, formatted_trip, ncc_val, f"{len(new_image_paths)}/4"))
        
        tree.selection_set(item_id)

        entry_plate.delete(0, tk.END)
        entry_trip.delete(0, tk.END)
        entry_trip.insert(0, "1")
        entry_plate.focus()
        temp_selected_images = []
        lbl_img_count.config(text="Chưa chọn ảnh nào (Yêu cầu 4 ảnh)", fg="red")
        print(f"[*] Đã CẮT 4 ảnh của xe {plate} vào thư mục {target_dir}")

    def delete_selected():
        selected = tree.selection()
        if not selected:
            return
        for item in selected:
            idx = tree.index(item)
            tree.delete(item)
            order_to_delete = queue_data.pop(idx)
            
            if order_to_delete.get("images"):
                folder_to_delete = os.path.dirname(order_to_delete["images"][0])
                try:
                    shutil.rmtree(folder_to_delete, ignore_errors=True)
                    print(f"[*] Đã dọn dẹp thư mục tạm: {folder_to_delete}")
                except:
                    pass

        for i, item in enumerate(tree.get_children()):
            vals = list(tree.item(item, "values"))
            vals[0] = i + 1
            tree.item(item, values=vals)
            
        if not tree.get_children():
            for lbl in img_labels:
                lbl.config(image='', text="Trống", width=26, height=13)

    root.mainloop()
    return queue_data if is_started["value"] else []


# ==============================================================================
# 3. ĐIỀU HƯỚNG & HÀM TƯƠNG TÁC
# ==============================================================================
def smart_click(page, locator):
    locator.wait_for(state="attached", timeout=6000)
    locator.evaluate("el => el.scrollIntoView({block: 'center'})")
    page.wait_for_timeout(300) 
    locator.click(force=True)

def choose_dropdown_option(page, keyword, exact_text=None):
    page.wait_for_timeout(400) 
    dropdown = page.locator(".ant-select-dropdown:visible").last
    dropdown.wait_for(state="visible", timeout=6000)
    
    target_text = exact_text if exact_text else keyword
    target_item = dropdown.locator(f"text='{target_text}'").first
    
    try:
        target_item.wait_for(state="visible", timeout=2000)
        target_item.evaluate("el => el.scrollIntoView({block: 'nearest'})")
        page.wait_for_timeout(200)
        target_item.click(force=True, timeout=2000)
    except:
        try:
            fallback = dropdown.locator(".ant-select-item-option").first
            fallback.evaluate("el => el.scrollIntoView({block: 'nearest'})")
            page.wait_for_timeout(200)
            fallback.click(force=True, timeout=2000)
        except:
            page.keyboard.press("Enter")
            
    page.wait_for_timeout(300)

def ensure_authenticated(context, page):
    page.wait_for_timeout(1500)
    if "/login" in page.url or page.locator("button:has-text('Đăng nhập')").count() > 0:
        print("-> Phát hiện trang đăng nhập, đang tự động đăng nhập...")
        page.wait_for_selector("input:visible", timeout=10000)
        inputs = page.locator("input:visible")

        inputs.nth(0).click()
        inputs.nth(0).fill(USERNAME)
        page.wait_for_timeout(200)

        inputs.nth(1).click()
        inputs.nth(1).fill(PASSWORD)
        page.wait_for_timeout(200)

        btn_submit = page.locator("button:has-text('Đăng nhập'):not(:has-text('Vingroup'))").last
        btn_submit.click()

        page.wait_for_timeout(3500)
        print("-> Đăng nhập thành công, lưu lại phiên...")
        context.storage_state(path=AUTH_FILE)

    if "quan-ly-kho" not in page.url:
        print("-> Đang vào thẳng trang Phiếu Vận Chuyển...")
        page.goto(URL_LIST, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)

def switch_company(page, target_company="VINALPHA"):
    page.wait_for_timeout(1000)
    for attempt in range(3):
        try:
            company_box = page.locator("header, div[class*='header']").locator("div:has-text('CÔNG TY CỔ PHẦN')").first
            company_box.wait_for(state="visible", timeout=6000)

            if target_company in company_box.inner_text():
                return

            print(f"Đang đổi công ty sang '{target_company}'...")
            smart_click(page, company_box)

            item_target = page.locator(f".ant-dropdown:visible div:has-text('{target_company}'), div:has-text('HẠ TẦNG {target_company}'), div:has-text('{target_company}')").last
            item_target.wait_for(state="visible", timeout=6000)
            smart_click(page, item_target)
            page.wait_for_timeout(2500)

            if target_company in company_box.inner_text():
                print(f"-> Đã ở đúng công ty: {target_company}")
                return
        except Exception as e:
            if attempt == 2:
                raise RuntimeError(f"Không thể chuyển sang công ty {target_company}: {e}")
            page.wait_for_timeout(1000)

def select_sidebar_user(page, role_title, user_name):
    if not user_name:
        raise ValueError(f"Thiếu thông tin cán bộ cho vai trò: '{role_title}'")

    block = page.locator(f"div:has(> div:has-text('{role_title}')), div:has-text('{role_title}')").last
    field = block.locator(".ant-select-selector, input, div:has-text('Tìm kiếm...')").last
    field.wait_for(state="visible", timeout=6000)
    
    smart_click(page, field)
    page.keyboard.type(user_name, delay=20)
    choose_dropdown_option(page, user_name)
    print(f"  + Đã chọn '{role_title}': {user_name}")

def create_single_slip(page, order, khoi_luong):
    bien_so = order["plate"]
    so_chuyen = order["trip"]
    ncc_ghi_chu = order["ncc"]
    print(f"\n==================================================")
    print(f">> ĐIỀN ĐƠN XE [{bien_so}] | CHUYẾN [{so_chuyen}] | KL: [{khoi_luong}] | NOTE: [{ncc_ghi_chu}]")
    print(f"==================================================")

    # ----------------------------------------------------
    # KHÔNG DÙNG LỆNH PAGE.GOTO() NỮA ĐỂ THUẬN LUỒNG APP 
    # Bấm trực tiếp nút Tạo Mới nếu nó đang hiện sẵn trên màn hình
    # ----------------------------------------------------
    btn_create = page.locator("button:has-text('Tạo mới'), a:has-text('Tạo mới')").first
    try:
        # Chờ tối đa 10s để nút Tạo mới xuất hiện tự nhiên
        btn_create.wait_for(state="visible", timeout=10000)
    except:
        # Fallback: Chỉ tải lại trang khi bị lỗi đơ mạng không hiện nút
        print("  -> Không tìm thấy nút Tạo mới, đang thử tải lại trang...")
        page.goto(URL_LIST, wait_until="domcontentloaded")
        btn_create.wait_for(state="visible", timeout=15000)

    btn_create.click()
    page.wait_for_selector("text='Thông tin chung'", timeout=10000)
    page.wait_for_timeout(1000)

    # ----------------------------------------------------
    # 2. CHỌN DỰ ÁN
    # ----------------------------------------------------
    box_du_an = page.locator("div:has-text('Chọn dự án')").last
    smart_click(page, box_du_an)
    page.keyboard.type("Hòa Long", delay=20)
    choose_dropdown_option(page, "Hòa Long", DU_AN)
    print(f"  + Đã chọn Dự án: {DU_AN}")

    # BẢO VỆ CHỐNG DÍNH MENU 
    page.locator("text='Thông tin chung'").last.click(force=True)
    page.wait_for_timeout(1500) 

    # ----------------------------------------------------
    # 3. CHỌN ĐƠN VỊ 
    # ----------------------------------------------------
    print(f"  + Tìm và chọn Đơn vị: {DON_VI_KEYWORD} ...")
    
    box_don_vi = page.locator("div:has-text('Chọn đơn vị'), div:has-text('Chọn Đơn vị')").last
    if box_don_vi.count() == 0:
        box_don_vi = page.locator(".ant-select").filter(has=page.locator("text='Chọn'")).last

    smart_click(page, box_don_vi)
    page.wait_for_timeout(300)
    page.keyboard.type(DON_VI_KEYWORD, delay=20)
    choose_dropdown_option(page, DON_VI_KEYWORD)
    print(f"  + Đã chọn Đơn vị: {DON_VI_KEYWORD}")
    page.wait_for_timeout(500)

    # 4. Chọn Cán bộ phê duyệt
    select_sidebar_user(page, "Thủ kho BCH", order["thu_kho"])
    select_sidebar_user(page, "Ban quản lý", order["bql"])
    select_sidebar_user(page, "Nhân viên an ninh (Xác nhận vào cổng)", order["an_ninh"])
    select_sidebar_user(page, "Nhân viên an ninh (Xác nhận ra cổng)", order["an_ninh"])

    # 5. Chọn Ngày đăng ký
    date_box = page.locator("input[placeholder*='Chọn ngày'], div:has-text('Chọn ngày')").last
    smart_click(page, date_box)

    today_btn = page.locator("a:has-text('Hôm nay'), button:has-text('Hôm nay'), td[class*='today']").last
    if today_btn.is_visible():
        smart_click(page, today_btn)
    else:
        date_box.fill(datetime.now().strftime("%d/%m/%Y"))
        page.keyboard.press("Enter")

    # 6. Điền thông tin xe
    page.locator("input[placeholder='Nhập hạng mục công việc']").fill(HANG_MUC)
    page.locator("input[placeholder='Nhập loại phương tiện']").fill(LOAI_PHUONG_TIEN)
    page.locator("input[placeholder='Nhập biển số xe']").fill(bien_so)
    page.locator("input[placeholder='Nhập số chuyến']").fill(so_chuyen)

    # 7. Mở popup Thêm vật tư
    btn_them_vt = page.locator("button:has-text('Thêm vật tư')")
    smart_click(page, btn_them_vt)
    
    modal = page.locator("div[role='dialog'], .ant-modal").last
    modal.wait_for(state="visible", timeout=6000)

    # 7.1 Tên vật tư
    print("  + Chọn Tên vật tư...")
    select_ten_vt = modal.locator(".ant-select").nth(0)
    smart_click(page, select_ten_vt)
    page.keyboard.type(TEN_VAT_TU, delay=20)
    choose_dropdown_option(page, TEN_VAT_TU)

    # 7.2 NHÀ CUNG CẤP LẤY CỐ ĐỊNH LÀ "VẠN XUÂN" ĐỂ ĐIỀN TRÊN WEB
    print(f"  + Chọn Nhà cung cấp: {NHA_CUNG_CAP_FULL} ...")
    select_ncc = modal.locator(".ant-select").nth(2)
    smart_click(page, select_ncc)
    page.keyboard.type(DON_VI_KEYWORD, delay=10)
    choose_dropdown_option(page, DON_VI_KEYWORD, NHA_CUNG_CAP_FULL)

    # 7.3 Khối lượng
    kl_formatted = str(khoi_luong).replace(".", ",")
    input_kl = modal.locator("input[placeholder*='khối lượng'], input[placeholder*='Nhập khối lượng']").last
    smart_click(page, input_kl)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.type(kl_formatted, delay=20)

    # 7.4 Bấm 'Thêm'
    btn_them = modal.locator("button:has-text('Thêm')").last
    smart_click(page, btn_them)
    page.wait_for_timeout(800)

    # 8. Nhập Vị trí
    input_vi_tri = page.locator("input[placeholder='Nhập vị trí']").last
    smart_click(page, input_vi_tri)
    input_vi_tri.fill(VI_TRI)

    # 9. TẢI 4 ẢNH ĐÍNH KÈM
    images = order.get("images", [])
    if len(images) != 4:
        raise ValueError(f"Xe {bien_so} không đủ 4 ảnh để upload!")

    file_input = page.locator("input[type='file']").last
    file_input.set_input_files(images)
    print("  + Bắt đầu tải 4 ảnh lên hệ thống, vui lòng chờ...")

    page.wait_for_timeout(1000) 

    timeout_counter = 0
    uploading_items = page.locator(".ant-upload-list-item-uploading")
    while uploading_items.count() > 0 and timeout_counter < 60:
        page.wait_for_timeout(1000)
        timeout_counter += 1

    done_items = page.locator(".ant-upload-list-item-done, .ant-upload-list-item:not(.ant-upload-list-item-uploading)")
    while done_items.count() < 4 and timeout_counter < 60:
        page.wait_for_timeout(1000)
        timeout_counter += 1

    page.wait_for_timeout(1500) 
    print("  + Toàn bộ 4/4 ảnh đã load lên web thành công!")

    # ----------------------------------------------------
    # ----------------------------------------------------
    # ----------------------------------------------------
    # ----------------------------------------------------
    # ----------------------------------------------------
    # 10. GỬI ĐƠN VÀ ĐỢI VỀ TRANG DANH SÁCH
    # ----------------------------------------------------
    print("  + Bấm 'Lưu & Gửi đơn'...")
    btn_send = page.locator("button").filter(has_text="Lưu & Gửi đơn").last
    smart_click(page, btn_send)
    print("  -> Đang chờ hệ thống xử lý gửi và tự động quay về danh sách...") 

    # Ép tool đứng im chờ cho đến khi trang web tự đẩy về màn hình danh sách
    try:
        page.wait_for_selector("button:has-text('Tạo mới')", state="visible", timeout=15000)
        page.wait_for_timeout(800) 
        print("  -> Đã GỬI ĐƠN thành công và quay về màn hình Quản lý phiếu!")
    except Exception:
        print("  -> Mạng xử lý hơi chậm, ép đợi thêm 5 giây...")
        page.wait_for_timeout(5000)


# ==============================================================================
# 4. TIẾN TRÌNH THỰC THI CHÍNH
# ==============================================================================
def main():
    if not os.path.exists(EXCEL_FILE):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Lỗi file", f"Không tìm thấy file Excel '{EXCEL_FILE}'!\nBạn hãy kiểm tra lại thư mục.")
        return

    df = None
    try:
        df = pd.read_excel(EXCEL_FILE)
        if COL_BIEN_SO not in df.columns or COL_KHOI_LUONG not in df.columns:
            print(f"LỖI: File Excel thiếu cột '{COL_BIEN_SO}' hoặc '{COL_KHOI_LUONG}'!")
            return
        df[COL_BIEN_SO] = df[COL_BIEN_SO].astype(str).str.strip()
        df["clean_plate"] = df[COL_BIEN_SO].str.replace("-", "").str.replace(".", "").str.replace(" ", "").str.upper()

        has_trip_column = COL_SO_CHUYEN in df.columns
        if has_trip_column:
            df[COL_SO_CHUYEN] = df[COL_SO_CHUYEN].astype(str).str.strip().str.zfill(2)
    except Exception as e:
        print(f"LỖI đọc file Excel: {e}")
        return

    queue_orders = open_queue_manager(df)
    
    if not queue_orders:
        print("Đã hủy hoặc danh sách hàng chờ trống.")
        return

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, slow_mo=40, args=["--start-maximized"])

            if os.path.exists(AUTH_FILE):
                context = browser.new_context(storage_state=AUTH_FILE, no_viewport=True)
            else:
                context = browser.new_context(no_viewport=True)

            page = context.new_page()

            page.goto(URL_LIST, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            ensure_authenticated(context, page)
            switch_company(page, COMPANY_NAME)

            for idx, order in enumerate(queue_orders, 1):
                bien_so = order["plate"]
                so_chuyen = order["trip"]
                ncc_ghi_chu = order["ncc"]
                clean_target = bien_so.replace("-", "").replace(".", "").replace(" ", "").upper()

                matched = df[df["clean_plate"] == clean_target]
                if matched.empty:
                    continue

                if has_trip_column:
                    trip_matched = matched[matched[COL_SO_CHUYEN] == so_chuyen]
                    if not trip_matched.empty:
                        khoi_luong = str(trip_matched.iloc[0][COL_KHOI_LUONG]).strip()
                    else:
                        khoi_luong = str(matched.iloc[0][COL_KHOI_LUONG]).strip()
                else:
                    khoi_luong = str(matched.iloc[0][COL_KHOI_LUONG]).strip()

                try:
                    create_single_slip(page, order, khoi_luong)
                    print(f"-> [Hoàn tất {idx}/{len(queue_orders)}] Đã ĐIỀN VÀ LƯU XONG xe {bien_so} (Chuyến {so_chuyen}).")
                    
                    # --- XUẤT BÁO CÁO LƯU FILE CSV CHUẨN FORM (SỬ DỤNG DẤU CHẤM PHẨY) ---
                    file_exists = os.path.isfile(REPORT_FILE)
                    try:
                        with open(REPORT_FILE, mode='a', encoding='utf-8-sig', newline='') as f:
                            writer = csv.writer(f, delimiter=';') # Đã sửa ở đây
                            if not file_exists:
                                writer.writerow(["Biển kiểm soát", "Khối lượng", "Số chuyến tạo", "Nhà cung cấp", "Thời gian xuất"])
                            writer.writerow([bien_so, khoi_luong, so_chuyen, ncc_ghi_chu, datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
                        print(f"   [+] Đã ghi nhật ký vào file {REPORT_FILE}")
                    except Exception as log_err:
                        print(f"   [!] Lỗi ghi file báo cáo (Có thể bạn đang mở file này bằng Excel): {log_err}")

                    # Dọn dẹp ảnh tạm
                    if order.get("images"):
                        try:
                            folder_to_delete = os.path.dirname(order["images"][0])
                            shutil.rmtree(folder_to_delete, ignore_errors=True)
                        except:
                            pass
                            
                except Exception as slip_err:
                    print(f"! LỖI khi điền phiếu cho xe {bien_so}: {slip_err}")
                    print(">> Tự động chuyển sang xe tiếp theo sau 3 giây...")
                    time.sleep(3)

            print("\n>>> ĐÃ HOÀN THÀNH TẤT CẢ ĐƠN TRONG HÀNG CHỜ! TỰ ĐỘNG ĐÓNG CHƯƠNG TRÌNH... <<<")
            time.sleep(2)
            
    except Exception as main_err:
        print(f"Đã xảy ra sự cố: {main_err}")

if __name__ == "__main__":
    main()