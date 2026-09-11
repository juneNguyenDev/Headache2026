import os
import sys
import time
import shutil
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
AUTH_FILE = "auth_state.json"
EXCEL_FILE = "Khối lượng.xlsx"
TEMP_UPLOAD_DIR = "Temp_Upload"

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
# 2. GIAO DIỆN HÀNG CHỜ (GIỮ LẠI PREVIEW VÀ CUT ẢNH)
# ==============================================================================
def open_queue_manager():
    os.makedirs(TEMP_UPLOAD_DIR, exist_ok=True)
    queue_data = []

    root = tk.Tk()
    root.title("Bộ lập hàng chờ tạo phiếu vận chuyển (Bản Ổn Định + Tự dọn ảnh)")
    root.geometry("1000x820")
    root.resizable(False, False)

    font_lbl = ("Segoe UI", 10)
    font_entry = ("Segoe UI", 10)

    temp_selected_images = []

    frame_top = tk.Frame(root)
    frame_top.pack(fill="both", expand=True)

    frame_left = tk.Frame(frame_top, width=450)
    frame_left.pack(side="left", fill="both", padx=15, pady=15, expand=True)

    frame_right = tk.Frame(frame_top, width=500)
    frame_right.pack(side="right", fill="both", padx=15, pady=15, expand=True)

    tk.Label(frame_left, text="1. Biển số xe:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 2))
    entry_plate = tk.Entry(frame_left, font=font_entry)
    entry_plate.pack(fill="x", pady=2)
    entry_plate.focus()

    tk.Label(frame_left, text="2. Số chuyến (ví dụ: 1, 2, 3...):", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
    entry_trip = tk.Entry(frame_left, font=font_entry)
    entry_trip.insert(0, "1")
    entry_trip.pack(fill="x", pady=2)

    tk.Label(frame_left, text="3. Ảnh đính kèm (Bắt buộc đủ 4 ảnh):", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 2))
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

    frame_users = tk.LabelFrame(frame_left, text=" Cán bộ phê duyệt trong ca ", font=font_lbl)
    frame_users.pack(fill="x", pady=10)

    tk.Label(frame_users, text="Thủ kho BCH:").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    entry_thukho = tk.Entry(frame_users, font=font_entry, width=22)
    entry_thukho.insert(0, "Nguyễn Văn Nhật")
    entry_thukho.grid(row=0, column=1, padx=8, pady=3)

    tk.Label(frame_users, text="Ban quản lý:").grid(row=1, column=0, sticky="w", padx=8, pady=3)
    entry_bql = tk.Entry(frame_users, font=font_entry, width=22)
    entry_bql.insert(0, "Trịnh Văn Cần")
    entry_bql.grid(row=1, column=1, padx=8, pady=3)

    tk.Label(frame_users, text="An ninh:").grid(row=2, column=0, sticky="w", padx=8, pady=3)
    entry_anninh = tk.Entry(frame_users, font=font_entry, width=22)
    entry_anninh.insert(0, "Nguyễn Đình Tây")
    entry_anninh.grid(row=2, column=1, padx=8, pady=3)

    btn_add = tk.Button(frame_left, text="+ THÊM VÀO HÀNG CHỜ", font=("Segoe UI", 11, "bold"), bg="#1890ff", fg="white", command=lambda: add_to_queue(), height=2)
    btn_add.pack(fill="x", pady=15)

    tk.Label(frame_right, text="DANH SÁCH ĐƠN CHỜ XỬ LÝ:", font=("Segoe UI", 10, "bold")).pack(anchor="w")

    tree = ttk.Treeview(frame_right, columns=("STT", "BienSo", "Chuyen", "SoAnh"), show="headings", height=10)
    tree.heading("STT", text="STT")
    tree.heading("BienSo", text="Biển số")
    tree.heading("Chuyen", text="Chuyến")
    tree.heading("SoAnh", text="Ảnh")
    tree.column("STT", width=40, anchor="center")
    tree.column("BienSo", width=140, anchor="center")
    tree.column("Chuyen", width=70, anchor="center")
    tree.column("SoAnh", width=80, anchor="center")
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

    frame_preview = tk.LabelFrame(root, text=" Xem trước 4 ảnh của đơn đang chọn (Kích thước lớn) ", font=("Segoe UI", 10, "bold"))
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

        if not plate:
            messagebox.showerror("Lỗi", "Vui lòng nhập Biển số xe!")
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
            "images": new_image_paths,
            "thu_kho": entry_thukho.get().strip(),
            "bql": entry_bql.get().strip(),
            "an_ninh": entry_anninh.get().strip()
        }
        queue_data.append(order_item)

        stt = len(queue_data)
        item_id = tree.insert("", "end", values=(stt, plate, formatted_trip, f"{len(new_image_paths)}/4"))
        
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
# 3. ĐIỀU HƯỚNG & HÀM TƯƠNG TÁC (CHẬM & ỔN ĐỊNH TUYỆT ĐỐI)
# ==============================================================================
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

    # ĐIỀU HƯỚNG THẲNG ĐẾN TRANG PHIẾU VẬN CHUYỂN, BỎ QUA GIAO DIỆN HOME
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
            company_box.click(force=True)
            page.wait_for_timeout(800)

            item_target = page.locator(f".ant-dropdown:visible div:has-text('{target_company}'), div:has-text('HẠ TẦNG {target_company}'), div:has-text('{target_company}')").last
            item_target.wait_for(state="visible", timeout=6000)
            item_target.click(force=True)
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
    field.click(force=True)
    page.wait_for_timeout(400)

    page.keyboard.type(user_name, delay=40)
    page.wait_for_timeout(700)

    option = page.locator(f".ant-select-dropdown:visible .ant-select-item-option:has-text('{user_name}'), .ant-select-dropdown:visible div:has-text('{user_name}')").last
    option.wait_for(state="visible", timeout=6000)
    option.click(force=True)
    page.wait_for_timeout(400)
    print(f"  + Đã chọn '{role_title}': {user_name}")


def create_single_slip(page, order, khoi_luong):
    bien_so = order["plate"]
    so_chuyen = order["trip"]
    print(f"\n==================================================")
    print(f">> ĐIỀN ĐƠN XE [{bien_so}] | CHUYẾN [{so_chuyen}] | KL: [{khoi_luong}]")
    print(f"==================================================")

    page.goto(URL_LIST, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)

    # 1. Bấm '+ Tạo mới'
    btn_create = page.locator("button:has-text('Tạo mới'), a:has-text('Tạo mới')").first
    btn_create.wait_for(state="visible", timeout=15000)
    btn_create.click()
    page.wait_for_selector("text='Thông tin chung'", timeout=10000)
    page.wait_for_timeout(1000)

    # 2. Chọn Dự án
    box_du_an = page.locator("div:has-text('Chọn dự án')").last
    box_du_an.click(force=True)
    page.wait_for_timeout(500)

    page.keyboard.type("Hòa Long", delay=40)
    page.wait_for_timeout(600)

    item_du_an = page.locator(f"div:has-text('{DU_AN}'), .ant-select-item-option:has-text('Hòa Long')").last
    item_du_an.wait_for(state="visible", timeout=10000)
    item_du_an.click(force=True)
    print(f"  + Đã chọn Dự án: {DU_AN}")
    page.wait_for_timeout(2000)

    # 3. Chọn Đơn vị
    box_don_vi = page.locator(".ant-select").filter(has=page.locator("text='Chọn'")).last
    box_don_vi.wait_for(state="visible", timeout=6000)
    box_don_vi.click(force=True)
    page.wait_for_timeout(600)

    option_dv = page.locator(".ant-select-dropdown:visible").locator(f"text={DON_VI_KEYWORD}").first
    if option_dv.count() == 0:
        option_dv = page.locator(".ant-select-dropdown:visible .ant-select-item-option").first
    option_dv.wait_for(state="visible", timeout=6000)
    option_dv.click(force=True)
    print(f"  + Đã chọn Đơn vị: {DON_VI_KEYWORD}")
    page.wait_for_timeout(2500)

    # 4. Chọn Cán bộ phê duyệt
    select_sidebar_user(page, "Thủ kho BCH", order["thu_kho"])
    select_sidebar_user(page, "Ban quản lý", order["bql"])
    select_sidebar_user(page, "Nhân viên an ninh (Xác nhận vào cổng)", order["an_ninh"])
    select_sidebar_user(page, "Nhân viên an ninh (Xác nhận ra cổng)", order["an_ninh"])

    # 5. Chọn Ngày đăng ký (Ngày hiện tại)
    date_box = page.locator("input[placeholder*='Chọn ngày'], div:has-text('Chọn ngày')").last
    date_box.click(force=True)
    page.wait_for_timeout(400)

    today_btn = page.locator("a:has-text('Hôm nay'), button:has-text('Hôm nay'), td[class*='today']").last
    if today_btn.is_visible():
        today_btn.click(force=True)
    else:
        today_str = datetime.now().strftime("%d/%m/%Y")
        date_box.fill(today_str)
        page.keyboard.press("Enter")

    # 6. Điền thông tin xe & chuyến
    page.locator("input[placeholder='Nhập hạng mục công việc']").fill(HANG_MUC)
    page.locator("input[placeholder='Nhập loại phương tiện']").fill(LOAI_PHUONG_TIEN)
    page.locator("input[placeholder='Nhập biển số xe']").fill(bien_so)
    page.locator("input[placeholder='Nhập số chuyến']").fill(so_chuyen)

    # 7. Mở popup Thêm vật tư
    page.locator("button:has-text('Thêm vật tư')").click()
    modal = page.locator("div[role='dialog'], .ant-modal").last
    modal.wait_for(state="visible", timeout=6000)
    page.wait_for_timeout(800)

    # 7.1. TÊN VẬT TƯ: Dropdown thứ nhất (nth 0)
    print("  + Chọn Tên vật tư...")
    select_ten_vt = modal.locator(".ant-select").nth(0)
    select_ten_vt.wait_for(state="visible", timeout=5000)
    select_ten_vt.click(force=True)
    page.wait_for_timeout(400)

    page.keyboard.type(TEN_VAT_TU, delay=60)
    page.wait_for_timeout(800)
    item_vt = page.locator(".ant-select-dropdown:visible").locator(f"text='{TEN_VAT_TU}'").first
    item_vt.wait_for(state="visible", timeout=6000)
    item_vt.click(force=True)
    print(f"  -> Đã chọn Tên vật tư: {TEN_VAT_TU}")
    page.wait_for_timeout(800)

    # 7.2. NHÀ CUNG CẤP: Dropdown thứ ba (nth 2)
    print("  + Chọn Nhà cung cấp...")
    select_ncc = modal.locator(".ant-select").nth(2)
    select_ncc.wait_for(state="visible", timeout=5000)
    select_ncc.click(force=True)
    page.wait_for_timeout(400)

    page.keyboard.type(DON_VI_KEYWORD, delay=60)
    page.wait_for_timeout(800)
    item_ncc = page.locator(".ant-select-dropdown:visible").locator(f"text='{NHA_CUNG_CAP_FULL}'").first
    if item_ncc.count() == 0:
        item_ncc = page.locator(".ant-select-dropdown:visible").locator(f"text='{DON_VI_KEYWORD}'").first
    item_ncc.wait_for(state="visible", timeout=6000)
    item_ncc.click(force=True)
    print(f"  -> Đã chọn Nhà cung cấp: {NHA_CUNG_CAP_FULL}")
    page.wait_for_timeout(500)

    # 7.3. KHỐI LƯỢNG
    kl_formatted = str(khoi_luong).replace(".", ",")
    input_kl = modal.locator("input[placeholder*='khối lượng'], input[placeholder*='Nhập khối lượng']").last
    input_kl.click(force=True)
    page.wait_for_timeout(200)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.wait_for_timeout(150)
    page.keyboard.type(kl_formatted, delay=50)
    page.wait_for_timeout(600)

    # 7.4. Bấm 'Thêm'
    modal.locator("button:has-text('Thêm')").last.click(force=True)
    page.wait_for_timeout(1000)

    # 8. Nhập Vị trí
    input_vi_tri = page.locator("input[placeholder='Nhập vị trí']").last
    input_vi_tri.wait_for(state="visible", timeout=5000)
    input_vi_tri.fill(VI_TRI)
    page.wait_for_timeout(500)

    # 9. TẢI 4 ẢNH ĐÍNH KÈM & CHỜ UPLOAD XONG HOÀN TOÀN
    images = order.get("images", [])
    if len(images) != 4:
        raise ValueError(f"Xe {bien_so} không đủ 4 ảnh để upload!")

    file_input = page.locator("input[type='file']").last
    file_input.set_input_files(images)
    print("  + Bắt đầu tải 4 ảnh lên hệ thống, vui lòng chờ...")

    page.wait_for_timeout(1000)  # Đợi 1 giây để UI kích hoạt trạng thái upload

    timeout_counter = 0
    
    # 9.1: Chờ cho đến khi TẤT CẢ các trạng thái "đang tải" (uploading) biến mất
    uploading_items = page.locator(".ant-upload-list-item-uploading")
    while uploading_items.count() > 0 and timeout_counter < 60:
        page.wait_for_timeout(1000)
        timeout_counter += 1

    # 9.2: Chờ để đảm bảo đã có đủ 4 ảnh (trạng thái "done") xuất hiện trên web
    done_items = page.locator(".ant-upload-list-item-done, .ant-upload-list-item:not(.ant-upload-list-item-uploading)")
    while done_items.count() < 4 and timeout_counter < 60:
        page.wait_for_timeout(1000)
        timeout_counter += 1

    page.wait_for_timeout(1500)  # Khoảng chờ an toàn cuối cùng trước khi bấm nút
    print("  + Toàn bộ 4/4 ảnh đã load lên web thành công!")

    # 10. TỰ ĐỘNG BẤM LƯU LẠI
    print("  + Bấm 'Lưu lại'...")
    btn_save = page.locator("button").filter(has_text="Lưu lại").last
    btn_save.wait_for(state="visible", timeout=5000)
    btn_save.scroll_into_view_if_needed()
    btn_save.click()
    print("  -> Đang chờ hệ thống xử lý lưu dữ liệu...")
    page.wait_for_timeout(4000)  # Chờ 4 giây để API hoàn tất


# ==============================================================================
# 4. TIẾN TRÌNH THỰC THI CHÍNH
# ==============================================================================
def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"LỖI: Không tìm thấy file Excel '{EXCEL_FILE}'!")
        return

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

    # Mở giao diện lập hàng chờ
    queue_orders = open_queue_manager()
    if not queue_orders:
        print("Đã hủy hoặc danh sách hàng chờ trống.")
        return

    try:
        with sync_playwright() as p:
            # SỬ DỤNG SLOW_MO=40 VÀ ĐIỀU CHỈNH TRÌNH DUYỆT ĐỂ BẢO ĐẢM ỔN ĐỊNH
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

            # Xử lý tuần tự từng đơn tự động
            for idx, order in enumerate(queue_orders, 1):
                bien_so = order["plate"]
                so_chuyen = order["trip"]
                clean_target = bien_so.replace("-", "").replace(".", "").replace(" ", "").upper()

                matched = df[df["clean_plate"] == clean_target]
                if matched.empty:
                    print(f"! [Đơn {idx}/{len(queue_orders)}] Bỏ qua xe {bien_so}: Không tìm thấy biển số trong Excel.")
                    continue

                if has_trip_column:
                    trip_matched = matched[matched[COL_SO_CHUYEN] == so_chuyen]
                    if not trip_matched.empty:
                        khoi_luong = str(trip_matched.iloc[0][COL_KHOI_LUONG]).strip()
                    else:
                        print(f"! Không tìm thấy chuyến {so_chuyen} của xe {bien_so}, lấy dòng đầu tiên.")
                        khoi_luong = str(matched.iloc[0][COL_KHOI_LUONG]).strip()
                else:
                    khoi_luong = str(matched.iloc[0][COL_KHOI_LUONG]).strip()

                try:
                    create_single_slip(page, order, khoi_luong)
                    print(f"-> [Hoàn tất {idx}/{len(queue_orders)}] Đã ĐIỀN VÀ LƯU XONG xe {bien_so} (Chuyến {so_chuyen}).")
                    print(">> Tự động chuyển sang đơn tiếp theo...")
                    
                    # DỌN DẸP ẢNH TẠM (Chỉ chạy khi phiếu tạo thành công)
                    if order.get("images"):
                        try:
                            folder_to_delete = os.path.dirname(order["images"][0])
                            shutil.rmtree(folder_to_delete, ignore_errors=True)
                            print(f"   Đã dọn dẹp thư mục ảnh tạm của xe {bien_so}.")
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