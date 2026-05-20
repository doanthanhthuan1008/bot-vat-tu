# Bot Duyệt Vật Tư - Công ty CORNER

## Hướng dẫn deploy lên Railway.app

### Bước 1: Chuẩn bị file
Tải về máy các file sau:
- bot.py
- requirements.txt
- Procfile
- mau_vat_tu.xlsx (file mẫu CORNER của bạn)

### Bước 2: Tạo tài khoản GitHub
1. Vào github.com → Sign up (miễn phí)
2. Tạo repository mới → đặt tên "bot-vat-tu"
3. Upload tất cả 4 file lên repository đó

### Bước 3: Deploy lên Railway
1. Vào railway.app → Login with GitHub
2. Bấm "New Project" → "Deploy from GitHub repo"
3. Chọn repo "bot-vat-tu"
4. Railway sẽ tự detect và deploy

### Bước 4: Cài biến môi trường
Trong Railway → tab "Variables" → thêm:
- BOT_TOKEN = token bot của bạn
- SEP_CHAT_ID = 1585175827
- NHOM_VAT_TU_ID = -5152630862
- GOOGLE_SHEETS_ID = (ID Google Sheets để lưu báo cáo)

### Bước 5: Khởi động
Bấm "Deploy" → Railway sẽ tự chạy bot 24/7

## Cách sử dụng

### Nhân viên:
1. Tìm @duyetvattu_bot trên Telegram
2. Gõ /deuxuat
3. Trả lời từng câu hỏi của bot
4. Bấm "Thêm vật tư nữa" nếu cần
5. Bấm "Xong, gửi duyệt" khi hoàn tất

### Sếp duyệt:
1. Nhận file Excel + tóm tắt từ bot
2. Bấm ✅ DUYỆT hoặc ❌ TỪ CHỐI
3. Bot tự động gửi vào nhóm vật tư nếu duyệt

### Báo cáo tháng:
- Mở Google Sheets → lọc theo tháng
- Xuất ra Excel để báo cáo
