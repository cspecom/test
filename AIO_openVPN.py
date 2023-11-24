import importlib
import os
import subprocess

import requests
import hashlib
import time

# Danh sách các thư viện cần kiểm tra và cài đặt
required_libraries = [
    ("pyautogui", "pyautogui"),
    ("pywinauto", "pywinauto"),
    ("requests", "requests"),
    ("numpy", "numpy"),
    ("mouse", "mouse"),
    ("psutil", "psutil"),
    ("win32crypt", "pypiwin32"),
    ("Crypto", "pycryptodome"),
    ("html2text", "html2text")
]

for module_name, package_name in required_libraries:
    spec = importlib.util.find_spec(module_name)
    if spec is None:
        print(f"Thư viện {package_name} chưa được cài đặt, đang tiến hành cài đặt...")
        try:
            subprocess.check_call(["pip", "install", package_name])
            print(f"Thư viện {package_name} đã được cài đặt thành công.")
            time.sleep(10)
        except Exception as e:
            print(f"Lỗi trong quá trình cài đặt thư viện {package_name}. Vui lòng cài đặt thủ công.")
            print(e)

# Lấy đường dẫn tuyệt đối của thư mục đang thực thi chương trình
current_dir = os.path.abspath(os.getcwd())

# Đường dẫn tương đối của file.py
file_path = "AIO_openVPN.pyc"

# Kết hợp đường dẫn tuyệt đối của thư mục đang thực thi chương trình với đường dẫn tương đối của file.py
file_path = os.path.join(current_dir, file_path)

# Đường dẫn đến file.py trên GitHub
url = "https://raw.githubusercontent.com/cspecom/test/main/AIO_openVPN.pyc"

try:
    # Đọc nội dung file từ server
    response = requests.get(url)

    if response.status_code == 200:
        # Tạo một file mới với nội dung trên server
        with open(file_path + ".new", "wb") as f:
            f.write(response.content)

        # So sánh hash của nội dung file cũ và file mới
        if os.path.isfile(file_path):
            with open(file_path, "rb") as f1, open(file_path + ".new", "rb") as f2:
                hash1 = hashlib.sha256(f1.read()).hexdigest()
                hash2 = hashlib.sha256(f2.read()).hexdigest()

            # Nếu hai hash khác nhau, thực hiện cập nhật
            if hash1 != hash2:
                print("Đã có phiên bản mới, tiến hành cập nhật...")
                # Đổi tên file.py cũ thành tên mới
                timestamp = int(time.time())
                os.rename(file_path, file_path + "." + str(timestamp) + ".old")
                # Đổi tên file_new.py thành file.py
                os.rename(file_path + ".new", file_path)

                for filename in os.listdir(current_dir):
                    if filename.endswith(".old") or filename.endswith(".new") :
                        os.remove(os.path.join(current_dir, filename))

                # Thực thi lại file.py với nội dung mới
                print("Cập nhật phiên bản mới thành công")
                os.system("python " + file_path)

            else:
                print("Bạn đang chạy phiên bản mới nhất!")
                for filename in os.listdir(current_dir):
                    if filename.endswith(".old") or filename.endswith(".new") :
                        os.remove(os.path.join(current_dir, filename))

                os.system("python " + file_path)
        else:
            # File cũ không tồn tại, thực hiện cập nhật
            print("Đã có phiên bản mới, tiến hành cập nhật...")
            # Đổi tên file_new.py thành file.py
            os.rename(file_path + ".new", file_path)
            # Thực thi lại file.py với nội dung mới
            print("Cập nhật phiên bản mới thành công")
            for filename in os.listdir(current_dir):
                if filename.endswith(".old") or filename.endswith(".new"):
                    os.remove(os.path.join(current_dir, filename))

            os.system("python " + file_path)
    else:
        print("Không thể kết nối đến server, dừng cập nhật. Mã lỗi:", response.status_code)
        for filename in os.listdir(current_dir):
            if filename.endswith(".old") or filename.endswith(".new"):
                os.remove(os.path.join(current_dir, filename))

        os.system("python " + file_path)

except requests.exceptions.RequestException as e:
    print("Lỗi khi tải xuống file từ server, dừng cập nhật. Mã lỗi:", e)
    os.system("python " + file_path)

except IOError as e:
    print("Lỗi khi mở file, dừng cập nhật. Mã lỗi:", e)
    os.system("python " + file_path)

except OSError as e:
    print("Lỗi khi đổi tên file, dừng cập nhật. Mã lỗi:", e)
    os.system("python " + file_path)

except Exception as e:
    print("Lỗi không mong muốn xảy ra, dừng cập nhật. Mã lỗi:", e)
    os.system("python " + file_path)
