from playwright.sync_api import sync_playwright
import time

LOGIN_URL = "https://construction.v-office.vn/login"
AUTH_FILE = "auth_state.json"

print("=" * 60)
print("KHỞI TẠO VÀ LƯU PHIÊN ĐĂNG NHẬP")
print("=" * 60)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, args=["--start-maximized"])
    context = browser.new_context(no_viewport=True)
    page = context.new_page()

    page.goto(LOGIN_URL)
    
    print("\n>> Hãy đăng nhập tài khoản của bạn trên trình duyệt vừa mở.")
    print(">> Sau khi đăng nhập xong và màn hình đã vào được trang chính (hoặc Quản lý kho),")
    input(">> QUAY LẠI ĐÂY BẤM PHÍM ENTER ĐỂ TIẾN HÀNH LƯU PHIÊN: ")

    # Lưu đầy đủ cả Cookies và LocalStorage
    context.storage_state(path=AUTH_FILE)
    print(f"\n-> ĐÃ LƯU THÀNH CÔNG VÀO FILE: {AUTH_FILE}")
    print("-> Từ bây giờ file main.py sẽ không cần đăng nhập nữa!")
    
    browser.close() 