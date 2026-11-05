"""Xử lý mail Outlook (classic, qua COM): AI phân loại, báo Discord kèm file, nhắc lịch hẹn trước 1 ngày."""
import base64
import email
import email.policy
import email.utils
import imaplib
import json
import logging
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import requests
import win32com.client

BASE = Path(__file__).parent
CONFIG = json.loads((BASE / "config.json").read_text(encoding="utf-8"))
STATE_FILE = BASE / "state.json"
DOCS_DIR = BASE / "tai_lieu"
DOC_EXTS = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".ppt", ".pptx",
            ".txt", ".rtf", ".odt", ".ods", ".odp", ".zip", ".rar", ".7z"}
DISCORD_MAX_FILE = 10 * 1024 * 1024  # Discord giới hạn ~10MB/file cho webhook
COLORS = {"Công việc": 0x3B82F6, "Học tập": 0x22C55E, "Hóa đơn": 0xF59E0B,
          "Quảng cáo": 0x9CA3AF, "Xác nhận đăng nhập": 0x64748B, "Khác": 0x8B5CF6}

logging.basicConfig(
    filename=BASE / "log.txt", level=logging.INFO, encoding="utf-8",
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger()


# ---------- Discord ----------
def discord(text=None, embed=None, files=()):
    payload = {}
    if text:
        payload["content"] = text[:1900]
    if embed:
        payload["embeds"] = [embed]
    handles = [open(p, "rb") for p in files[:10]]
    try:
        r = requests.post(
            CONFIG["discord_webhook"],
            data={"payload_json": json.dumps(payload)},
            files={f"files[{i}]": (Path(p).name, h) for i, (p, h) in enumerate(zip(files, handles))},
            timeout=60,
        )
        r.raise_for_status()
    finally:
        for h in handles:
            h.close()


def windows_toast(title, body):
    """Hiện thông báo góc phải màn hình Windows (tắt bằng "windows_notify": false trong config)."""
    if not CONFIG.get("windows_notify", True):
        return
    try:
        from win11toast import notify
        notify(title[:100], body[:250], app_id="XuLiMail")
    except Exception as e:
        log.warning("Lỗi thông báo Windows: %s", e)


def mail_embed(account, item, a, cat, events, saved):
    details = a.get("details") or []
    fields = [{"name": "👤 Từ", "value": f"{item.SenderName} <{item.SenderEmailAddress}>"[:1024]}]
    if details:
        fields.append({"name": "📝 Chi tiết", "value": "\n".join(f"• {d}" for d in details)[:1024]})
    if events:
        fields.append({"name": "📅 Lịch hẹn (sẽ nhắc trước 1 ngày)", "value": "\n".join(
            f"• **{e['title']}** — {datetime.fromisoformat(e['start']):%H:%M %d/%m/%Y}"
            + (f" · 📍 {e['location']}" if e.get("location") else "") for e in events)[:1024]})
    if saved:
        fields.append({"name": f"📎 File đính kèm ({len(saved)})", "value": "\n".join(
            f"• {p.name}" + (" *(quá lớn, chỉ lưu trên máy)*" if p.stat().st_size > DISCORD_MAX_FILE else "")
            for p in saved)[:1024]})
    fields.append({"name": "📂 Nhánh", "value": cat, "inline": True})
    fields.append({"name": "🕒 Nhận lúc", "value": f"{local_dt(item.ReceivedTime):%H:%M %d/%m/%Y}", "inline": True})
    return {
        "author": {"name": f"📬 {account}"},
        "title": (item.Subject or "(không tiêu đề)")[:256],
        "description": f"### 📌 Nội dung chính\n> {a.get('main', '')}"[:4000],
        "color": COLORS.get(cat, 0x8B5CF6),
        "fields": fields,
        "footer": {"text": item.SenderEmailAddress or ""},
    }


# ---------- State (mọi thời gian đều là giờ máy, không kèm múi giờ) ----------
def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    since = (datetime.now() - timedelta(days=1)).isoformat(timespec="seconds")
    return {"last_check": since, "processed": [], "reminders": []}


def save_state(state):
    state["processed"] = state["processed"][-500:]
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------- Outlook ----------
def local_dt(t):
    # pywin32 trả về datetime gắn nhãn UTC nhưng thực chất là giờ máy -> bỏ tzinfo
    return datetime(t.year, t.month, t.day, t.hour, t.minute, t.second)


def get_inboxes():
    """Trả về [(tên tài khoản, Inbox)] cho mọi tài khoản trong Outlook (lọc theo config 'accounts' nếu có)."""
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    only = [x.lower() for x in CONFIG.get("accounts", [])]
    result = []
    for store in ns.Stores:
        if only and store.DisplayName.lower() not in only:
            continue
        if store.DisplayName.startswith("Online Archive"):
            continue  # hộp lưu trữ online: không nhận mail mới, hay lỗi mạng
        try:
            result.append((store.DisplayName, store.GetDefaultFolder(6)))  # 6 = Inbox
        except Exception:
            pass  # file lưu trữ PST, public folder... không có Inbox
    return result


def get_folders(inbox, names):
    existing = {f.Name: f for f in inbox.Folders}
    for n in names:
        if n not in existing:
            existing[n] = inbox.Folders.Add(n)
    return existing


def new_messages(inbox, since):
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # mới nhất trước
    result = []
    for item in items:
        if item.Class != 43:  # 43 = MailItem
            continue
        if local_dt(item.ReceivedTime) < since:
            break
        result.append(item)
    return list(reversed(result))  # cũ trước, mới sau


# ---------- Gmail (IMAP + mật khẩu ứng dụng) ----------
def imap_utf7(s):
    """Mã hóa tên thư mục kiểu IMAP (modified UTF-7) để dùng được tiếng Việt."""
    out, buf = [], []

    def flush():
        if buf:
            b64 = base64.b64encode("".join(buf).encode("utf-16-be")).decode().rstrip("=")
            out.append("&" + b64.replace("/", ",") + "-")
            buf.clear()
    for ch in s:
        if 0x20 <= ord(ch) <= 0x7E:
            flush()
            out.append("&-" if ch == "&" else ch)
        else:
            buf.append(ch)
    flush()
    return '"' + "".join(out) + '"'


class GmailAttachment:
    def __init__(self, part):
        self.part = part
        self.FileName = part.get_filename() or ""

    def SaveAsFile(self, path):
        Path(path).write_bytes(self.part.get_payload(decode=True) or b"")


class GmailItem:
    """Bọc 1 mail Gmail cho giống MailItem của Outlook để dùng chung code xử lý."""

    def __init__(self, imap, account, uid, raw):
        msg = email.message_from_bytes(raw, policy=email.policy.default)
        self.imap, self.uid = imap, uid
        self.SenderName, self.SenderEmailAddress = email.utils.parseaddr(str(msg["From"] or ""))
        self.SenderName = self.SenderName or self.SenderEmailAddress
        self.Subject = str(msg["Subject"] or "")
        self.ReceivedTime = email.utils.parsedate_to_datetime(msg["Date"]).astimezone()
        self.EntryID = f"gmail:{account}:{msg['Message-ID'] or uid.decode()}"
        body = msg.get_body(preferencelist=("plain", "html"))
        text = body.get_content() if body else ""
        if body and body.get_content_type() == "text/html":
            text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", text))
        self.Body = text
        self.Attachments = [GmailAttachment(p) for p in msg.iter_attachments()]

    def Move(self, label):
        self.imap.uid("MOVE", self.uid, imap_utf7(label))


def gmail_connect(acc):
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(acc["email"], acc["app_password"].replace(" ", ""))
    for cat in CONFIG["categories"]:
        imap.create(imap_utf7(cat))  # đã tồn tại thì Gmail báo lỗi nhẹ, bỏ qua
    imap.select("INBOX")
    return imap


def gmail_messages(imap, account, since, latest_only=False):
    _, data = imap.uid("SEARCH", None, f"SINCE {since:%d-%b-%Y}")
    uids = data[0].split()
    if latest_only:
        uids = uids[-1:] or imap.uid("SEARCH", None, "ALL")[1][0].split()[-1:]
    items = []
    for uid in uids:
        _, d = imap.uid("FETCH", uid, "(BODY.PEEK[])")  # PEEK: không đánh dấu đã đọc
        item = GmailItem(imap, account, uid, d[0][1])
        if latest_only or local_dt(item.ReceivedTime) >= since:
            items.append(item)
    return items


def save_attachments(item):
    """Lưu file tài liệu đính kèm vào tai_lieu/, tên có tiền tố ngày nhận. Bỏ qua ảnh chữ ký, v.v."""
    saved = []
    date = local_dt(item.ReceivedTime).strftime("%Y-%m-%d")
    for att in item.Attachments:
        name = att.FileName or ""
        if Path(name).suffix.lower() not in DOC_EXTS:
            continue
        DOCS_DIR.mkdir(exist_ok=True)
        path = DOCS_DIR / f"{date}_{name}"
        n = 1
        while path.exists():
            path = DOCS_DIR / f"{date}_{Path(name).stem} ({n}){Path(name).suffix}"
            n += 1
        try:
            att.SaveAsFile(str(path))
            saved.append(path)
            log.info("Lưu file %s", path.name)
        except Exception as e:
            log.exception("Lỗi lưu file %s: %s", name, e)
    return saved


# ---------- AI ----------
PROVIDER = CONFIG.get("provider", "gemini")


def ask_ai(prompt):
    if PROVIDER == "claude":
        import anthropic
        resp = anthropic.Anthropic(api_key=CONFIG["anthropic_api_key"]).messages.create(
            model=CONFIG.get("claude_model", "claude-haiku-5-5"), max_tokens=1000,
            messages=[{"role": "user", "content": prompt}],
        )
        return resp.content[0].text

    url = (f"https://generativelanguage.googleapis.com/v1beta/models/"
           f"{CONFIG.get('gemini_model', 'gemini-3.5-flash-lite')}:generateContent")
    body = {"contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"responseMimeType": "application/json"}}
    for _ in range(4):  # gói free giới hạn số request/phút -> chờ rồi thử lại
        r = requests.post(url, params={"key": CONFIG["gemini_api_key"]}, json=body, timeout=60)
        if r.status_code != 429:
            break
        time.sleep(30)
    r.raise_for_status()
    parts = r.json()["candidates"][0]["content"]["parts"]
    return "".join(p.get("text", "") for p in parts if not p.get("thought"))


def analyze(item):
    now = datetime.now().strftime("%A %Y-%m-%d %H:%M")
    attachments = ", ".join(a.FileName for a in item.Attachments) or "không có"
    prompt = f"""Bây giờ là {now}. Phân tích email sau, trả lời bằng tiếng Việt.

- category: chọn đúng 1 nhóm trong: {", ".join(CONFIG["categories"])}
  ("Xác nhận đăng nhập" = mã OTP, mã xác minh, cảnh báo/thông báo đăng nhập, đặt lại mật khẩu, xác nhận tài khoản)
- main: nội dung chính của mail trong 1-2 câu ngắn gọn (mail muốn gì / thông báo gì)
- details: 2-5 ý chi tiết quan trọng (hạn chót, yêu cầu, số tiền, người liên quan, link...). Ngắn gọn.
- events: lịch hẹn/cuộc họp/deadline có ngày giờ cụ thể trong tương lai (mảng rỗng nếu không có)

Trả về CHỈ JSON:
{{"category": "...", "main": "...", "details": ["...", "..."],
  "events": [{{"title": "...", "start": "YYYY-MM-DDTHH:MM", "location": "..."}}]}}

Người gửi: {item.SenderName} <{item.SenderEmailAddress}>
Ngày nhận: {local_dt(item.ReceivedTime):%Y-%m-%d %H:%M}
Tiêu đề: {item.Subject}
File đính kèm: {attachments}
Nội dung:
{(item.Body or "")[:6000]}"""
    text = ask_ai(prompt).strip()
    text = text[text.find("{"): text.rfind("}") + 1]
    return json.loads(text)


# ---------- Main ----------
def handle(account, item, state):
    """Phân tích 1 mail, gửi thông báo Discord. Trả về nhánh đã chọn."""
    a = analyze(item)
    cat = a.get("category") if a.get("category") in CONFIG["categories"] else CONFIG["categories"][-1]
    if cat in CONFIG.get("silent_categories", []):
        log.info("[%s] [%s] (im lặng) %s", account, cat, item.Subject)
        return cat

    events = []
    for ev in a.get("events") or []:
        try:
            start = datetime.fromisoformat(ev["start"])
        except Exception:
            continue
        if start > datetime.now():
            events.append({"title": ev.get("title") or item.Subject, "start": start.isoformat(),
                           "location": ev.get("location", ""), "subject": item.Subject})
    state["reminders"].extend(events)

    saved = save_attachments(item)
    uploads = [p for p in saved if p.stat().st_size <= DISCORD_MAX_FILE]
    discord(embed=mail_embed(account, item, a, cat, events, saved), files=uploads)
    extra = "".join([f"\n📅 {len(events)} lịch hẹn" if events else "",
                     f"\n📎 {len(saved)} file" if saved else ""])
    windows_toast(f"✉️ {item.SenderName} · {cat}",
                  f"{item.Subject}\n📌 {a.get('main', '')}{extra}")
    log.info("[%s] [%s] %s", account, cat, item.Subject)
    return cat


def process_mail(state):
    run_start = datetime.now().isoformat(timespec="seconds")
    since = datetime.fromisoformat(state["last_check"])
    try:
        inboxes = get_inboxes() if CONFIG.get("use_outlook", True) else []
    except Exception as e:
        log.exception("Lỗi mở Outlook: %s", e)
        inboxes = []
    for account, inbox in inboxes:
        try:
            folders = get_folders(inbox, CONFIG["categories"])
            for item in new_messages(inbox, since):
                if item.EntryID in state["processed"]:
                    continue
                try:
                    cat = handle(account, item, state)
                except Exception as e:
                    log.exception("Lỗi xử lý %s: %s", item.Subject, e)
                    continue
                item.Move(folders[cat])
                state["processed"].append(item.EntryID)
        except Exception as e:
            log.exception("Lỗi hộp thư %s: %s", account, e)

    for acc in CONFIG.get("gmail_accounts", []):
        account = acc["email"]
        try:
            imap = gmail_connect(acc)
            for item in gmail_messages(imap, account, since):
                if item.EntryID in state["processed"]:
                    continue
                try:
                    cat = handle(account, item, state)
                except Exception as e:
                    log.exception("Lỗi xử lý %s: %s", item.Subject, e)
                    continue
                item.Move(cat)
                state["processed"].append(item.EntryID)
            imap.logout()
        except Exception as e:
            log.exception("Lỗi Gmail %s: %s", account, e)
    state["last_check"] = run_start


def check_reminders(state):
    now = datetime.now()
    keep = []
    for r in state["reminders"]:
        start = datetime.fromisoformat(r["start"])
        if start < now:
            continue  # đã qua, bỏ
        last = datetime.fromisoformat(r["last_reminded"]) if r.get("last_reminded") else None
        every = timedelta(hours=CONFIG.get("remind_every_hours", 5))
        # Trong 24h trước giờ hẹn: nhắc lần đầu, rồi cứ mỗi `every` nhắc lại
        if start - now <= timedelta(days=1) and (last is None or now - last >= every):
            left = start - now
            left_text = f"còn {left.seconds // 3600 + left.days * 24} giờ {left.seconds % 3600 // 60} phút"
            discord(embed={
                "title": f"⏰ Nhắc lịch: {r['title']}"[:256],
                "description": f"### 🕒 {start:%H:%M %d/%m/%Y} ({left_text})"
                               + (f"\n📍 {r['location']}" if r.get("location") else ""),
                "color": 0xEF4444,
                "footer": {"text": f"Từ mail: {r['subject']}"[:2048]},
            })
            windows_toast(f"⏰ Nhắc lịch: {r['title']}",
                          f"🕒 {start:%H:%M %d/%m/%Y} ({left_text})"
                          + (f"\n📍 {r['location']}" if r.get("location") else ""))
            r["last_reminded"] = now.isoformat(timespec="seconds")
        keep.append(r)
    state["reminders"] = keep


def test_latest():
    """Chạy thử trên mail mới nhất của mỗi hộp thư: gửi thông báo nhưng KHÔNG chuyển thư mục, KHÔNG lưu state."""
    for acc in CONFIG.get("gmail_accounts", []):
        imap = gmail_connect(acc)
        for item in gmail_messages(imap, acc["email"], datetime.now() - timedelta(days=1), latest_only=True):
            print(f"[{acc['email']}] Mail:", item.Subject)
            print("Đã gửi thông báo. Nhánh:", handle(acc["email"], item, {"reminders": []}))
        imap.logout()
    for account, inbox in (get_inboxes() if CONFIG.get("use_outlook", True) else []):
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        item = next((i for i in items if i.Class == 43), None)
        if item is None:
            continue
        print(f"[{account}] Mail:", item.Subject)
        cat = handle(account, item, {"reminders": []})
        print("Đã gửi thông báo. Nhánh:", cat)


def main():
    state = load_state()
    try:
        process_mail(state)
    except Exception as e:
        log.exception("Lỗi: %s", e)
    check_reminders(state)
    save_state(state)


if __name__ == "__main__":
    if "--test-discord" in sys.argv:
        discord("✅ Kết nối Discord OK")
    elif "--test-latest" in sys.argv:
        test_latest()
    else:
        main()
