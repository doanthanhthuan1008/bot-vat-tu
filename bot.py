import os
import json
import logging
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters, ConversationHandler
)
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
import gspread
from google.oauth2.service_account import Credentials
import copy

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ====== CẤU HÌNH ======
BOT_TOKEN = os.environ.get("BOT_TOKEN", "8835628733:AAGF1OAt55CA9JyrvhKdesw7a_qqb6R4ez8")
SEP_CHAT_ID = int(os.environ.get("SEP_CHAT_ID", "1585175827"))
NHOM_VAT_TU_ID = int(os.environ.get("NHOM_VAT_TU_ID", "-5152630862"))
GOOGLE_SHEETS_ID = os.environ.get("GOOGLE_SHEETS_ID", "")
MAU_FILE = "mau_vat_tu.xlsx"

# ====== TRẠNG THÁI HỘI THOẠI ======
(
    CHON_HANH_DONG,
    NHAP_BO_PHAN,
    NHAP_MA_VAT_TU,
    NHAP_TEN_VAT_TU,
    NHAP_DVT,
    NHAP_NHU_CAU,
    NHAP_TON_KHO,
    NHAP_SO_LUONG,
    NHAP_XUAT_XU,
    NHAP_MUC_DICH,
    NHAP_NGAY_GIAO,
    NHAP_GHI_CHU,
    THEM_VAT_TU,
    XAC_NHAN,
) = range(14)

# ====== LƯU DỮ LIỆU TẠM ======
user_data_store = {}


def tao_phieu_excel(don_hang: dict) -> str:
    wb = load_workbook(MAU_FILE)
    ws = wb.active

    bo_phan = don_hang.get("bo_phan", "")
    nguoi_de_nghi = don_hang.get("nguoi_de_nghi", "")
    danh_sach_vat_tu = don_hang.get("vat_tu", [])
    ngay_tao = datetime.now()

    # Điền phòng/bộ phận vào dòng 12
    for cell in ws[12]:
        if cell.value and "Phòng/Bộ phận" in str(cell.value):
            cell.value = f"Đề nghị cung cấp cho Phòng/Bộ phận/Đơn vị: {bo_phan} một số vật tư, cụ thể như sau:"
            break

    # Điền ngày tháng vào dòng 31
    ws['H31'] = f"Đà Nẵng, ngày {ngay_tao.day} tháng {ngay_tao.month} năm {ngay_tao.year}"

    # Điền tên người đề nghị
    ws['I32'] = nguoi_de_nghi

    # Điền dữ liệu vật tư từ dòng 16 (STT 1)
    hang_bat_dau = 16
    for i, vt in enumerate(danh_sach_vat_tu[:10]):
        row = hang_bat_dau + i
        ws.cell(row=row, column=2).value = vt.get("ma_vat_tu", "")
        ws.cell(row=row, column=3).value = vt.get("ten_vat_tu", "")
        ws.cell(row=row, column=4).value = vt.get("dvt", "")
        ws.cell(row=row, column=5).value = vt.get("nhu_cau", "")
        ws.cell(row=row, column=6).value = vt.get("ton_kho", "")
        ws.cell(row=row, column=7).value = vt.get("so_luong_mua", "")
        ws.cell(row=row, column=8).value = vt.get("xuat_xu", "")
        ws.cell(row=row, column=9).value = vt.get("muc_dich", "")
        ws.cell(row=row, column=10).value = vt.get("ngay_giao", "")
        ws.cell(row=row, column=11).value = vt.get("ghi_chu", "")

    ten_file = f"phieu_vat_tu_{don_hang['id']}.xlsx"
    wb.save(ten_file)
    return ten_file


def tao_tin_nhan_tom_tat(don_hang: dict) -> str:
    vat_tu_list = don_hang.get("vat_tu", [])
    lines = [
        f"📋 *PHIẾU YÊU CẦU VẬT TƯ*",
        f"👤 Người đề nghị: {don_hang.get('nguoi_de_nghi', '')}",
        f"🏢 Bộ phận: {don_hang.get('bo_phan', '')}",
        f"📅 Ngày: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        f"\n*Chi tiết vật tư ({len(vat_tu_list)} mặt hàng):*",
    ]
    for i, vt in enumerate(vat_tu_list, 1):
        lines.append(
            f"\n*{i}. {vt['ten_vat_tu']}*"
            f"\n   • Mã: {vt.get('ma_vat_tu', '-')}"
            f"\n   • ĐVT: {vt.get('dvt', '-')}"
            f"\n   • Số lượng cần mua: {vt.get('so_luong_mua', '-')}"
            f"\n   • Mục đích: {vt.get('muc_dich', '-')}"
            f"\n   • Ngày giao: {vt.get('ngay_giao', '-')}"
        )
    return "\n".join(lines)


async def luu_google_sheets(don_hang: dict, trang_thai: str):
    if not GOOGLE_SHEETS_ID:
        return
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(GOOGLE_SHEETS_ID).sheet1
        ngay = datetime.now().strftime("%d/%m/%Y %H:%M")
        for vt in don_hang.get("vat_tu", []):
            sheet.append_row([
                ngay,
                don_hang.get("nguoi_de_nghi", ""),
                don_hang.get("bo_phan", ""),
                vt.get("ma_vat_tu", ""),
                vt.get("ten_vat_tu", ""),
                vt.get("dvt", ""),
                vt.get("so_luong_mua", ""),
                vt.get("muc_dich", ""),
                vt.get("ngay_giao", ""),
                trang_thai,
            ])
    except Exception as e:
        logger.error(f"Lỗi Google Sheets: {e}")


# ====== HANDLERS ======

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Xin chào! Tôi là *Bot Duyệt Vật Tư* của Công ty CORNER.\n\n"
        "Gõ /deuxuat để bắt đầu đề xuất vật tư.",
        parse_mode="Markdown"
    )


async def bat_dau_de_xuat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id
    user_data_store[user_id] = {
        "id": f"{user_id}_{int(datetime.now().timestamp())}",
        "nguoi_de_nghi": f"{user.first_name} {user.last_name or ''}".strip(),
        "bo_phan": "",
        "vat_tu": [],
        "vat_tu_hien_tai": {}
    }
    await update.message.reply_text(
        "📝 *BẮT ĐẦU TẠO PHIẾU YÊU CẦU VẬT TƯ*\n\n"
        "Bạn thuộc Phòng/Bộ phận nào?\n"
        "_(Ví dụ: Phòng Kỹ thuật, Bộ phận Sản xuất...)_",
        parse_mode="Markdown"
    )
    return NHAP_BO_PHAN


async def nhap_bo_phan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_store[user_id]["bo_phan"] = update.message.text
    await update.message.reply_text(
        "✅ Tiếp theo, nhập *Mã vật tư* (nếu không có thì gõ dấu `-`)",
        parse_mode="Markdown"
    )
    return NHAP_MA_VAT_TU


async def nhap_ma_vat_tu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    ma = update.message.text.strip()
    user_data_store[user_id]["vat_tu_hien_tai"]["ma_vat_tu"] = "" if ma == "-" else ma
    await update.message.reply_text(
        "📦 Nhập *Tên vật tư* (kèm thông số kỹ thuật, màu sắc, loại nếu có)\n"
        "_(Ví dụ: Ốc vít M6x20 inox, Dây điện 2.5mm màu đỏ...)_",
        parse_mode="Markdown"
    )
    return NHAP_TEN_VAT_TU


async def nhap_ten_vat_tu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_store[user_id]["vat_tu_hien_tai"]["ten_vat_tu"] = update.message.text
    await update.message.reply_text("📏 Nhập *Đơn vị tính* (cái, kg, m, hộp, cuộn...)", parse_mode="Markdown")
    return NHAP_DVT


async def nhap_dvt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_store[user_id]["vat_tu_hien_tai"]["dvt"] = update.message.text
    await update.message.reply_text("🔢 Nhập *Nhu cầu sử dụng* (số lượng dự kiến dùng)", parse_mode="Markdown")
    return NHAP_NHU_CAU


async def nhap_nhu_cau(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_store[user_id]["vat_tu_hien_tai"]["nhu_cau"] = update.message.text
    await update.message.reply_text(
        "🏭 Nhập *Tồn kho hiện tại* (nếu không biết thì gõ `0`)",
        parse_mode="Markdown"
    )
    return NHAP_TON_KHO


async def nhap_ton_kho(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_store[user_id]["vat_tu_hien_tai"]["ton_kho"] = update.message.text
    await update.message.reply_text("🛒 Nhập *Số lượng cần mua*", parse_mode="Markdown")
    return NHAP_SO_LUONG


async def nhap_so_luong(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_store[user_id]["vat_tu_hien_tai"]["so_luong_mua"] = update.message.text
    await update.message.reply_text(
        "🌍 Nhập *Xuất xứ* (nếu không rõ thì gõ dấu `-`)\n_(Ví dụ: Việt Nam, Nhật, Trung Quốc...)_",
        parse_mode="Markdown"
    )
    return NHAP_XUAT_XU


async def nhap_xuat_xu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    xuat_xu = update.message.text.strip()
    user_data_store[user_id]["vat_tu_hien_tai"]["xuat_xu"] = "" if xuat_xu == "-" else xuat_xu
    await update.message.reply_text("🎯 Nhập *Mục đích sử dụng*", parse_mode="Markdown")
    return NHAP_MUC_DICH


async def nhap_muc_dich(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_store[user_id]["vat_tu_hien_tai"]["muc_dich"] = update.message.text
    await update.message.reply_text(
        "📅 Nhập *Ngày giao hàng mong muốn*\n_(Ví dụ: 30/06/2025 hoặc gõ `-` nếu chưa xác định)_",
        parse_mode="Markdown"
    )
    return NHAP_NGAY_GIAO


async def nhap_ngay_giao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    ngay = update.message.text.strip()
    user_data_store[user_id]["vat_tu_hien_tai"]["ngay_giao"] = "" if ngay == "-" else ngay
    await update.message.reply_text(
        "📝 Nhập *Ghi chú* (nếu không có thì gõ dấu `-`)",
        parse_mode="Markdown"
    )
    return NHAP_GHI_CHU


async def nhap_ghi_chu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    ghi_chu = update.message.text.strip()
    user_data_store[user_id]["vat_tu_hien_tai"]["ghi_chu"] = "" if ghi_chu == "-" else ghi_chu

    vt = user_data_store[user_id]["vat_tu_hien_tai"]
    user_data_store[user_id]["vat_tu"].append(copy.deepcopy(vt))
    user_data_store[user_id]["vat_tu_hien_tai"] = {}

    so_luong = len(user_data_store[user_id]["vat_tu"])
    keyboard = [
        [InlineKeyboardButton("➕ Thêm vật tư nữa", callback_data="them_vat_tu")],
        [InlineKeyboardButton("✅ Xong, gửi duyệt", callback_data="gui_duyet")],
    ]
    await update.message.reply_text(
        f"✅ Đã thêm vật tư thứ *{so_luong}*!\n\nBạn muốn làm gì tiếp theo?",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown"
    )
    return THEM_VAT_TU


async def xu_ly_them_vat_tu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "them_vat_tu":
        await query.message.reply_text(
            "➕ Nhập *Mã vật tư* tiếp theo (gõ `-` nếu không có)",
            parse_mode="Markdown"
        )
        return NHAP_MA_VAT_TU

    elif query.data == "gui_duyet":
        don_hang = user_data_store[user_id]
        tom_tat = tao_tin_nhan_tom_tat(don_hang)
        keyboard = [
            [InlineKeyboardButton("✅ Xác nhận gửi", callback_data="xac_nhan_gui")],
            [InlineKeyboardButton("❌ Hủy bỏ", callback_data="huy_bo")],
        ]
        await query.message.reply_text(
            f"📋 *XEM LẠI PHIẾU TRƯỚC KHI GỬI:*\n\n{tom_tat}",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )
        return XAC_NHAN


async def xu_ly_xac_nhan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "huy_bo":
        await query.message.reply_text("❌ Đã hủy phiếu. Gõ /deuxuat để tạo lại.")
        user_data_store.pop(user_id, None)
        return ConversationHandler.END

    don_hang = user_data_store[user_id]
    await query.message.reply_text("⏳ Đang tạo phiếu và gửi cho sếp duyệt...")

    try:
        ten_file = tao_phieu_excel(don_hang)
        tom_tat = tao_tin_nhan_tom_tat(don_hang)
        keyboard = [
            [
                InlineKeyboardButton("✅ DUYỆT", callback_data=f"duyet_{don_hang['id']}"),
                InlineKeyboardButton("❌ TỪ CHỐI", callback_data=f"tuchoi_{don_hang['id']}"),
            ]
        ]

        with open(ten_file, "rb") as f:
            await context.bot.send_document(
                chat_id=SEP_CHAT_ID,
                document=f,
                filename=ten_file,
                caption=f"{tom_tat}\n\n👆 Vui lòng duyệt hoặc từ chối phiếu này.",
                reply_markup=InlineKeyboardMarkup(keyboard),
                parse_mode="Markdown"
            )

        user_data_store[f"don_{don_hang['id']}"] = {
            "don_hang": don_hang,
            "nguoi_gui_id": user_id,
            "file": ten_file
        }

        await query.message.reply_text(
            "✅ *Phiếu đã được gửi cho sếp duyệt!*\n"
            "Bạn sẽ nhận thông báo khi sếp phê duyệt.",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(f"Lỗi gửi phiếu: {e}")
        await query.message.reply_text("❌ Có lỗi xảy ra. Vui lòng thử lại sau.")

    user_data_store.pop(user_id, None)
    return ConversationHandler.END


async def xu_ly_duyet(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    data = query.data
    if data.startswith("duyet_"):
        don_id = data.replace("duyet_", "")
        trang_thai = "duyet"
    else:
        don_id = data.replace("tuchoi_", "")
        trang_thai = "tuchoi"

    key = f"don_{don_id}"
    if key not in user_data_store:
        await query.message.reply_text("⚠️ Không tìm thấy phiếu này.")
        return

    info = user_data_store[key]
    don_hang = info["don_hang"]
    nguoi_gui_id = info["nguoi_gui_id"]
    ten_file = info["file"]

    if trang_thai == "duyet":
        tom_tat = tao_tin_nhan_tom_tat(don_hang)
        with open(ten_file, "rb") as f:
            await context.bot.send_document(
                chat_id=NHOM_VAT_TU_ID,
                document=f,
                filename=ten_file,
                caption=f"✅ *PHIẾU ĐÃ ĐƯỢC DUYỆT*\n\n{tom_tat}\n\n👉 Vui lòng tiến hành mua vật tư.",
                parse_mode="Markdown"
            )
        await context.bot.send_message(
            chat_id=nguoi_gui_id,
            text="✅ *Phiếu yêu cầu vật tư của bạn đã được DUYỆT!*\nBộ phận vật tư sẽ tiến hành mua hàng.",
            parse_mode="Markdown"
        )
        await query.message.edit_caption(
            caption=query.message.caption + "\n\n✅ *ĐÃ DUYỆT*",
            parse_mode="Markdown"
        )
        await luu_google_sheets(don_hang, "Đã duyệt")
    else:
        await context.bot.send_message(
            chat_id=nguoi_gui_id,
            text="❌ *Phiếu yêu cầu vật tư của bạn đã bị TỪ CHỐI.*\nVui lòng liên hệ sếp để biết thêm chi tiết.",
            parse_mode="Markdown"
        )
        await query.message.edit_caption(
            caption=query.message.caption + "\n\n❌ *ĐÃ TỪ CHỐI*",
            parse_mode="Markdown"
        )
        await luu_google_sheets(don_hang, "Từ chối")

    user_data_store.pop(key, None)


async def huy_cuoc_tro_chuyen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data_store.pop(user_id, None)
    await update.message.reply_text("❌ Đã hủy. Gõ /deuxuat để bắt đầu lại.")
    return ConversationHandler.END


def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("deuxuat", bat_dau_de_xuat)],
        states={
            NHAP_BO_PHAN: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_bo_phan)],
            NHAP_MA_VAT_TU: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_ma_vat_tu)],
            NHAP_TEN_VAT_TU: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_ten_vat_tu)],
            NHAP_DVT: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_dvt)],
            NHAP_NHU_CAU: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_nhu_cau)],
            NHAP_TON_KHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_ton_kho)],
            NHAP_SO_LUONG: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_so_luong)],
            NHAP_XUAT_XU: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_xuat_xu)],
            NHAP_MUC_DICH: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_muc_dich)],
            NHAP_NGAY_GIAO: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_ngay_giao)],
            NHAP_GHI_CHU: [MessageHandler(filters.TEXT & ~filters.COMMAND, nhap_ghi_chu)],
            THEM_VAT_TU: [CallbackQueryHandler(xu_ly_them_vat_tu)],
            XAC_NHAN: [CallbackQueryHandler(xu_ly_xac_nhan)],
        },
        fallbacks=[CommandHandler("huy", huy_cuoc_tro_chuyen)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler)
    app.add_handler(CallbackQueryHandler(xu_ly_duyet, pattern="^(duyet_|tuchoi_)"))

    logger.info("Bot đang chạy...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
