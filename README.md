# XuLiMail

Script Python chạy nền trên Windows: đọc mail mới từ Outlook classic và Gmail, dùng AI phân loại và tóm tắt, gửi thông báo qua Discord, rồi nhắc lịch hẹn tìm thấy trong mail.

## Tính năng

- **Phân loại:** chuyển mỗi mail vào thư mục (Outlook) hoặc nhãn (Gmail) theo nhánh trong `categories`. Thư mục và nhãn còn thiếu sẽ được tạo tự động.
- **Thông báo:** gửi người gửi, nội dung chính và các ý quan trọng lên Discord, kèm thông báo Windows.
- **Tệp đính kèm:** lưu tài liệu vào `tai_lieu/` và gửi lên Discord. File trên 10 MB chỉ được lưu trên máy.
- **Nhắc lịch:** trong 24 giờ trước giờ hẹn, nhắc lần đầu rồi lặp lại sau mỗi `remind_every_hours` giờ.
- **Nhánh im lặng:** mail thuộc `silent_categories` vẫn được chuyển vào đúng nhánh, nhưng không gửi thông báo và không lưu tệp đính kèm.

## Yêu cầu

- Windows và Python 3
- Discord webhook
- Khóa API Gemini hoặc Anthropic
- Ít nhất một hộp thư:
  - Outlook classic đã đăng nhập trên máy, hoặc
  - Gmail đã bật mật khẩu ứng dụng

## Cài đặt

```powershell
pip install -r requirements.txt
```

Tạo file `config.json` cùng thư mục với `main.py`:

```json
{
  "provider": "gemini",
  "gemini_api_key": "<khóa Gemini>",
  "gemini_model": "gemini-3.5-flash-lite",
  "anthropic_api_key": "",
  "claude_model": "claude-haiku-5-5",
  "discord_webhook": "<URL Discord webhook>",
  "use_outlook": true,
  "gmail_accounts": [
    {"email": "<địa chỉ Gmail>", "app_password": "<mật khẩu ứng dụng>"}
  ],
  "categories": ["Công việc", "Học tập", "Hóa đơn", "Quảng cáo", "Xác nhận đăng nhập", "Khác"],
  "remind_every_hours": 5,
  "silent_categories": ["Xác nhận đăng nhập"]
}
```

| Trường | Ý nghĩa |
|---|---|
| `provider` | `gemini` hoặc `claude` |
| `use_outlook` | `false` nếu không dùng Outlook |
| `gmail_accounts` | Danh sách tài khoản Gmail, để `[]` nếu không dùng |
| `categories` | Các nhánh phân loại. Nhánh cuối là nhánh dự phòng khi AI trả về tên không hợp lệ |
| `accounts` | Không bắt buộc. Chỉ xử lý các hộp thư Outlook có tên trong danh sách |
| `windows_notify` | Không bắt buộc. `false` để tắt thông báo Windows |

`config.json` chứa khóa và mật khẩu, đã nằm trong `.gitignore`. Không commit file này.

## Chạy

```powershell
python main.py                  # xử lý mail mới và nhắc lịch một lần
python main.py --test-discord   # kiểm tra Discord webhook
python main.py --test-latest    # chạy thử trên mail mới nhất của mỗi hộp thư
```

`--test-latest` có gửi thông báo và lưu tệp đính kèm, nhưng không chuyển mail và không ghi `state.json`.

Lần chạy đầu tiên quét mail của 1 ngày gần nhất. Các lần sau chỉ quét từ lần chạy trước.

## Chạy tự động

Tạo tác vụ Windows chạy `main.py` mỗi 5 phút:

```powershell
powershell -ExecutionPolicy Bypass -File install_task.ps1
```

Xóa tác vụ:

```powershell
schtasks /delete /tn XuLiMail /f
```

## Dữ liệu và quyền riêng tư

- Người gửi, tiêu đề, tên tệp đính kèm và tối đa 6000 ký tự đầu của nội dung mail được gửi tới nhà cung cấp AI đã chọn.
- Bản tóm tắt và tệp đính kèm được gửi lên kênh Discord của webhook.
- Các file sau chỉ nằm trên máy và không được đưa lên Git:

| File | Nội dung |
|---|---|
| `config.json` | Khóa API, webhook, mật khẩu ứng dụng |
| `state.json` | Mail đã xử lý, lịch nhắc |
| `log.txt` | Nhật ký chạy, có tiêu đề mail |
| `tai_lieu/` | Tệp đính kèm đã lưu |
