import os
import logging
import re
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import calendar

# Load environment variables dari file .env
load_dotenv()

# Setup logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Konfigurasi
TELEGRAM_BOT_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
WORKSHEET_NAME = os.getenv('WORKSHEET_NAME', 'Pengeluaran')  # Default 'Pengeluaran'
CREDENTIALS_FILE = "service-account-key.json"

# Header constants (sama seperti yang kamu gunakan)
H_TANGGAL = "📅 Tanggal"
H_WAKTU = "⏰ Waktu"
H_KATEGORI = "🏷️ Kategori"
H_DESKRIPSI = "📝 Deskripsi"
H_JUMLAH = "💰 Jumlah (Rp)"
H_USER_ID = "👤 User ID"
H_STATUS = "📊 Status"

# Untuk summary worksheet header row = baris 3 (A3..)
SUMMARY_HEADER_ROW = 3
MAIN_HEADER_ROW = 1

# Validasi environment variables
if not TELEGRAM_BOT_TOKEN or not SPREADSHEET_ID:
    logger.error("❌ TELEGRAM_BOT_TOKEN dan SPREADSHEET_ID harus diisi di file .env")
    print("❌ Error: Pastikan file .env berisi:")
    print("TELEGRAM_BOT_TOKEN=your_bot_token_here")
    print("SPREADSHEET_ID=your_spreadsheet_id_here")
    exit(1)

logger.info(f"✅ Bot token loaded: {TELEGRAM_BOT_TOKEN[:10]}...")
logger.info(f"✅ Spreadsheet ID loaded: {SPREADSHEET_ID[:10]}...")

def to_int(value):
    """Konversi aman menjadi int meskipun value mengandung 'Rp', titik, koma, spasi, dsb."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(value)
    if value is None:
        return 0
    s = str(value).strip()
    # Hapus semua karakter non-digit kecuali minus
    s = re.sub(r'[^0-9\-]', '', s)
    if s in ('', '-'):
        return 0
    try:
        return int(s)
    except ValueError:
        return 0

class ExpenseBot:
    def __init__(self):
        self.gc = None
        self.worksheet = None
        self.summary_worksheet = None
        self.setup_google_sheets()

    def setup_google_sheets(self):
        """Setup koneksi ke Google Sheets"""
        try:
            scope = [
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.file",
                "https://www.googleapis.com/auth/drive"
            ]

            if not os.path.exists(CREDENTIALS_FILE):
                logger.error(f"❌ File {CREDENTIALS_FILE} tidak ditemukan!")
                raise FileNotFoundError(f"File {CREDENTIALS_FILE} tidak ada")

            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
            self.gc = gspread.authorize(creds)

            spreadsheet = self.gc.open_by_key(SPREADSHEET_ID)

            # Setup main worksheet
            try:
                self.worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
                logger.info(f"✅ Worksheet '{WORKSHEET_NAME}' ditemukan")
            except gspread.WorksheetNotFound:
                logger.info(f"📝 Membuat worksheet baru: '{WORKSHEET_NAME}'")
                self.worksheet = spreadsheet.add_worksheet(WORKSHEET_NAME, 1000, 7)
                self.setup_main_worksheet_format()

            # Setup summary worksheet
            try:
                self.summary_worksheet = spreadsheet.worksheet('Ringkasan Bulanan')
                logger.info("✅ Worksheet 'Ringkasan Bulanan' ditemukan")
            except gspread.WorksheetNotFound:
                logger.info("📝 Membuat worksheet 'Ringkasan Bulanan'")
                self.summary_worksheet = spreadsheet.add_worksheet('Ringkasan Bulanan', 50, 10)
                self.setup_summary_worksheet_format()

            logger.info(f"✅ Google Sheets berhasil terhubung ke: {spreadsheet.title}")

        except FileNotFoundError as e:
            logger.error(f"❌ File credentials tidak ditemukan: {e}")
            raise e
        except Exception as e:
            logger.error(f"❌ Error setup Google Sheets: {e}")
            raise e

    def setup_main_worksheet_format(self):
        """Setup format untuk worksheet utama"""
        try:
            headers = [H_TANGGAL, H_WAKTU, H_KATEGORI, H_DESKRIPSI, H_JUMLAH, H_USER_ID, H_STATUS]

            # Hanya append header sekali
            existing = self.worksheet.get_all_values()
            if not existing or len(existing) == 0:
                self.worksheet.append_row(headers)

            # Format header row (A1:G1)
            header_format = {
                "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.8},
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}
                },
                "horizontalAlignment": "CENTER"
            }
            self.worksheet.format("A1:G1", header_format)

            # Set column widths - per request API fields must include "fields"
            requests = [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": self.worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 1
                        },
                        "properties": {"pixelSize": 120}
                    },
                    "fields": "pixelSize"
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": self.worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 1,
                            "endIndex": 2
                        },
                        "properties": {"pixelSize": 100}
                    },
                    "fields": "pixelSize"
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": self.worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 2,
                            "endIndex": 3
                        },
                        "properties": {"pixelSize": 120}
                    },
                    "fields": "pixelSize"
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": self.worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 3,
                            "endIndex": 4
                        },
                        "properties": {"pixelSize": 250}
                    },
                    "fields": "pixelSize"
                },
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": self.worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 4,
                            "endIndex": 5
                        },
                        "properties": {"pixelSize": 150}
                    },
                    "fields": "pixelSize"
                }
            ]

            spreadsheet = self.gc.open_by_key(SPREADSHEET_ID)
            spreadsheet.batch_update({"requests": requests})

            logger.info("✅ Format worksheet utama berhasil disetup")

        except Exception as e:
            logger.error(f"❌ Error setup format worksheet utama: {e}")

    def setup_summary_worksheet_format(self):
        """Setup format untuk worksheet ringkasan"""
        try:
            # Title
            self.summary_worksheet.update('A1', '📊 RINGKASAN PENGELUARAN BULANAN 📊')

            # Headers untuk summary di baris 3
            headers = ["📅 Bulan-Tahun", "🍽️ Makan", "🚗 Transport", "🛒 Belanja", "🎮 Hiburan", "💊 Kesehatan", "📦 Lainnya", "💰 Total"]

            # Pastikan header hanya ditambahkan jika belum ada
            vals = self.summary_worksheet.get_all_values()
            if len(vals) < SUMMARY_HEADER_ROW or not any(vals[SUMMARY_HEADER_ROW-1]):
                # gunakan table_range agar berada di baris 3
                self.summary_worksheet.append_row(headers, table_range='A3')

            # Format title
            title_format = {
                "backgroundColor": {"red": 0.8, "green": 0.2, "blue": 0.2},
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "fontSize": 16
                },
                "horizontalAlignment": "CENTER"
            }
            self.summary_worksheet.format("A1:H1", title_format)
            self.summary_worksheet.merge_cells('A1:H1')

            # Format header row (A3:H3)
            header_format = {
                "backgroundColor": {"red": 0.2, "green": 0.6, "blue": 0.3},
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}
                },
                "horizontalAlignment": "CENTER"
            }
            self.summary_worksheet.format("A3:H3", header_format)

            logger.info("✅ Format worksheet ringkasan berhasil disetup")

        except Exception as e:
            logger.error(f"❌ Error setup format worksheet ringkasan: {e}")

    def parse_expense_message(self, message_text):
        """Parse pesan dengan format: kategori jumlah deskripsi"""
        pattern = r'^(\w+)\s+(\d+)\s+(.+)$'
        match = re.match(pattern, message_text.strip(), re.IGNORECASE)

        if match:
            kategori = match.group(1).lower()
            try:
                jumlah = int(match.group(2))
                deskripsi = match.group(3)
                return kategori, jumlah, deskripsi
            except ValueError:
                return None, None, None
        return None, None, None

    def add_expense_to_sheet(self, user_id, kategori, jumlah, deskripsi):
        """Tambahkan pengeluaran ke Google Sheets"""
        try:
            now = datetime.now()
            tanggal = now.strftime("%Y-%m-%d")
            waktu = now.strftime("%H:%M:%S")
            status = "✅ Berhasil"

            row_data = [tanggal, waktu, kategori.title(), deskripsi, jumlah, user_id, status]
            self.worksheet.append_row(row_data)

            # Update monthly summary
            self.update_monthly_summary()

            # Apply alternating row colors
            self.apply_row_formatting()

            logger.info(f"✅ Data ditambahkan: {kategori} - Rp{jumlah} - {deskripsi}")
            return True
        except Exception as e:
            logger.error(f"❌ Error menambahkan ke sheet: {e}")
            return False

    def apply_row_formatting(self):
        """Apply alternating row colors and number formatting"""
        try:
            records = self.worksheet.get_all_records(head=MAIN_HEADER_ROW, value_render_option='UNFORMATTED_VALUE')
            total_rows = len(records) + MAIN_HEADER_ROW  # + header rows

            # Alternating row colors (data mulai baris 2)
            for i in range(MAIN_HEADER_ROW + 1, total_rows + 1):
                if i % 2 == 0:
                    row_format = {"backgroundColor": {"red": 0.95, "green": 0.98, "blue": 1.0}}
                else:
                    row_format = {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}
                self.worksheet.format(f"A{i}:G{i}", row_format)

            # Format number column (column E) with currency formatting
            number_format = {
                "numberFormat": {
                    "type": "CURRENCY",
                    "pattern": "Rp #,##0"
                }
            }
            if total_rows >= 2:
                self.worksheet.format(f"E2:E{total_rows}", number_format)

        except Exception as e:
            logger.error(f"❌ Error formatting rows: {e}")

    def update_monthly_summary(self):
        """Update monthly summary in separate worksheet"""
        try:
            current_month_str = datetime.now().strftime("%Y-%m")    # format YYYY-MM untuk startswith
            current_month_name = datetime.now().strftime("%B %Y")  # "August 2025"

            records = self.worksheet.get_all_records(head=MAIN_HEADER_ROW, value_render_option='UNFORMATTED_VALUE')

            # Prepare monthly aggregation
            categories = ['makan', 'transport', 'belanja', 'hiburan', 'kesehatan', 'lainnya']
            monthly_data = {c: 0 for c in categories}

            for record in records:
                try:
                    raw_date = record.get(H_TANGGAL)
                    if not raw_date:
                        continue
                    # ensure string date 'YYYY-MM-DD' or convert if number. We'll stringify.
                    record_date_str = str(raw_date)
                    # Accept both 'YYYY-MM-DD' or datelike object -> compare startswith YYYY-MM
                    if record_date_str.startswith(current_month_str):
                        category = str(record.get(H_KATEGORI, "")).lower()
                        amount = to_int(record.get(H_JUMLAH))
                        if category in monthly_data:
                            monthly_data[category] += amount
                        else:
                            monthly_data['lainnya'] += amount
                except Exception:
                    continue

            total = sum(monthly_data.values())

            # Baca summary worksheet records (header ada di baris SUMMARY_HEADER_ROW)
            summary_records = self.summary_worksheet.get_all_records(head=SUMMARY_HEADER_ROW, value_render_option='UNFORMATTED_VALUE')
            existing_row = None
            for i, rec in enumerate(summary_records):
                if rec.get('📅 Bulan-Tahun') == current_month_name:
                    # data rows start pada baris SUMMARY_HEADER_ROW + 1 (misal header 3 -> data mulai 4)
                    existing_row = i + SUMMARY_HEADER_ROW + 1
                    break

            row_data = [
                current_month_name,
                monthly_data['makan'],
                monthly_data['transport'],
                monthly_data['belanja'],
                monthly_data['hiburan'],
                monthly_data['kesehatan'],
                monthly_data['lainnya'],
                total
            ]

            if existing_row:
                # update cell by cell
                for col, value in enumerate(row_data, start=1):
                    self.summary_worksheet.update_cell(existing_row, col, value)
            else:
                # append new row (row will be appended under the last)
                self.summary_worksheet.append_row(row_data)

            # format summary sheet
            self.format_summary_worksheet()

            logger.info(f"✅ Monthly summary updated for {current_month_name}")

        except Exception as e:
            logger.error(f"❌ Error updating monthly summary: {e}")

    def format_summary_worksheet(self):
        """Format the summary worksheet"""
        try:
            records = self.summary_worksheet.get_all_records(head=SUMMARY_HEADER_ROW, value_render_option='UNFORMATTED_VALUE')
            total_rows = len(records) + SUMMARY_HEADER_ROW  # header row included

            # Format currency columns (B..H), data rows mulai di baris SUMMARY_HEADER_ROW+1
            if total_rows >= SUMMARY_HEADER_ROW + 1:
                currency_format = {
                    "numberFormat": {
                        "type": "CURRENCY",
                        "pattern": "Rp #,##0"
                    }
                }
                self.summary_worksheet.format(f"B{SUMMARY_HEADER_ROW+1}:H{total_rows}", currency_format)

                # Highlight total column
                total_format = {
                    "backgroundColor": {"red": 1.0, "green": 0.9, "blue": 0.6},
                    "textFormat": {"bold": True}
                }
                self.summary_worksheet.format(f"H{SUMMARY_HEADER_ROW+1}:H{total_rows}", total_format)

        except Exception as e:
            logger.error(f"❌ Error formatting summary worksheet: {e}")

    def get_today_summary(self, user_id):
        """Dapatkan ringkasan pengeluaran hari ini"""
        try:
            today = datetime.now().strftime("%Y-%m-%d")
            records = self.worksheet.get_all_records(head=MAIN_HEADER_ROW, value_render_option='UNFORMATTED_VALUE')

            today_expenses = [
                record for record in records
                if str(record.get(H_TANGGAL, "")) == today and str(record.get(H_USER_ID, "")) == str(user_id)
            ]

            if not today_expenses:
                return "📊 Belum ada pengeluaran hari ini."

            total_today = sum(to_int(r.get(H_JUMLAH)) for r in today_expenses)
            count_today = len(today_expenses)

            summary = f"📊 *RINGKASAN PENGELUARAN HARI INI*\n"
            summary += f"📅 Tanggal: {datetime.now().strftime('%d %B %Y')}\n\n"
            summary += f"📝 Jumlah transaksi: {count_today}\n"
            summary += f"💰 Total pengeluaran: Rp {total_today:,}\n\n"

            # Categories today
            categories_today = {}
            for expense in today_expenses:
                cat = str(expense.get(H_KATEGORI, "")).lower()
                amount = to_int(expense.get(H_JUMLAH))
                categories_today[cat] = categories_today.get(cat, 0) + amount

            summary += "*📋 PENGELUARAN PER KATEGORI HARI INI:*\n"
            category_icons = {
                'makan': '🍽️',
                'transport': '🚗',
                'belanja': '🛒',
                'hiburan': '🎮',
                'kesehatan': '💊'
            }

            for cat, amount in categories_today.items():
                icon = category_icons.get(cat, '📦')
                summary += f"{icon} {cat.title()}: Rp {amount:,}\n"

            return summary

        except Exception as e:
            logger.error(f"❌ Error mendapatkan ringkasan: {e}")
            return "❌ Error mendapatkan ringkasan pengeluaran."

    def get_monthly_summary(self, user_id):
        """Dapatkan ringkasan pengeluaran bulan ini (khusus user tertentu)"""
        try:
            current_month = datetime.now().strftime("%Y-%m")  # YYYY-MM
            records = self.worksheet.get_all_records(head=MAIN_HEADER_ROW, value_render_option='UNFORMATTED_VALUE')

            monthly_expenses = [
                record for record in records
                if str(record.get(H_TANGGAL, "")).startswith(current_month) and str(record.get(H_USER_ID, "")) == str(user_id)
            ]

            if not monthly_expenses:
                return "📊 Belum ada pengeluaran bulan ini."

            total_monthly = sum(to_int(r.get(H_JUMLAH)) for r in monthly_expenses)
            count_monthly = len(monthly_expenses)

            # Calculate daily average & projection
            current_date = datetime.now()
            days_in_month = calendar.monthrange(current_date.year, current_date.month)[1]
            current_day = current_date.day
            daily_average = total_monthly / current_day if current_day > 0 else 0

            summary = f"📊 *RINGKASAN PENGELUARAN BULAN INI*\n"
            summary += f"📅 Bulan: {datetime.now().strftime('%B %Y')}\n\n"
            summary += f"📝 Total transaksi: {count_monthly}\n"
            summary += f"💰 Total pengeluaran: Rp {total_monthly:,}\n"
            summary += f"📊 Rata-rata per hari: Rp {daily_average:,.0f}\n\n"

            # Categories monthly
            categories_monthly = {}
            for expense in monthly_expenses:
                cat = str(expense.get(H_KATEGORI, "")).lower()
                amount = to_int(expense.get(H_JUMLAH))
                categories_monthly[cat] = categories_monthly.get(cat, 0) + amount

            summary += "*📋 PENGELUARAN PER KATEGORI BULAN INI:*\n"
            category_icons = {
                'makan': '🍽️',
                'transport': '🚗',
                'belanja': '🛒',
                'hiburan': '🎮',
                'kesehatan': '💊'
            }

            for cat, amount in sorted(categories_monthly.items(), key=lambda x: x[1], reverse=True):
                percentage = (amount / total_monthly * 100) if total_monthly > 0 else 0
                icon = category_icons.get(cat, '📦')
                summary += f"{icon} {cat.title()}: Rp {amount:,} ({percentage:.1f}%)\n"

            return summary

        except Exception as e:
            logger.error(f"❌ Error mendapatkan ringkasan bulanan: {e}")
            return "❌ Error mendapatkan ringkasan pengeluaran bulanan."

# Inisialisasi bot
logger.info("🔧 Menginisialisasi bot...")
expense_bot = ExpenseBot()
logger.info("✅ Bot berhasil diinisialisasi!")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /start"""
    user_name = update.message.from_user.first_name or "User"
    welcome_message = f"""
🤖 *Halo {user_name}! Selamat datang di Bot Pencatat Pengeluaran!*
    """
    await update.message.reply_text(welcome_message, parse_mode='Markdown')
    logger.info(f"👋 User {update.message.from_user.id} memulai bot")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /help"""
    help_text = (
        "📖 *PANDUAN PENGGUNAAN BOT*\n\n"
        "Klik salah satu tombol di bawah untuk mengirim perintah:\n\n"
        "_/start_ — untuk memulai bot\n"
        "_/help_ — untuk melihat panduan\n"
        "_/ringkasan_ — untuk melihat ringkasan harian dan bulanan\n\n"
        "Atau ketik langsung salah satu perintah di atas."
    )

    # Keyboard dengan tombol yang akan otomatis mengirim pesan ketika diklik
    keyboard = ReplyKeyboardMarkup(
        [["/start", "/help", "/ringkasan"]],
        resize_keyboard=True,
        one_time_keyboard=False  # ubah ke True jika ingin keyboard hilang setelah klik
    )

    await update.message.reply_text(help_text, parse_mode='Markdown', reply_markup=keyboard)

async def summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    logger.info(f"📊 User {user_id} meminta ringkasan")
    loading_msg = await update.message.reply_text("📊 Sedang mengambil data ringkasan...")
    try:
        daily_summary = expense_bot.get_today_summary(user_id)
        monthly_summary = expense_bot.get_monthly_summary(user_id)
        combined_summary = f"{daily_summary}\n\n{'='*40}\n\n{monthly_summary}"
        await loading_msg.edit_text(combined_summary, parse_mode='Markdown')
    except Exception as e:
        logger.error(f"❌ Error menampilkan ringkasan: {e}")
        await loading_msg.edit_text("❌ Gagal mengambil ringkasan. Silakan coba lagi.")

async def handle_expense(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    message_text = update.message.text
    logger.info(f"💬 User {user_id} mengirim: {message_text}")
    kategori, jumlah, deskripsi = expense_bot.parse_expense_message(message_text)
    if kategori is None:
        await update.message.reply_text("❌ Format pesan salah. Ketik /help untuk panduan.", parse_mode='Markdown')
        return
    loading_msg = await update.message.reply_text("💾 Menyimpan data dan memperbarui spreadsheet...")
    success = expense_bot.add_expense_to_sheet(user_id, kategori, jumlah, deskripsi)
    if success:
        await loading_msg.edit_text("✅ Pengeluaran berhasil dicatat!", parse_mode='Markdown')
    else:
        await loading_msg.edit_text("❌ Gagal menyimpan pengeluaran ke Google Sheets.", parse_mode='Markdown')

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.error(f"Update {update} caused error {context.error}")

def main():
    try:
        if not TELEGRAM_BOT_TOKEN:
            print("❌ TELEGRAM_BOT_TOKEN tidak ditemukan di file .env")
            return
        application = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
        application.add_handler(CommandHandler("start", start))
        application.add_handler(CommandHandler("help", help_command))
        application.add_handler(CommandHandler("ringkasan", summary))
        application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_expense))
        application.add_error_handler(error_handler)

        print("🤖 Bot berjalan...")
        application.run_polling(
            allowed_updates=Update.ALL_TYPES,
            drop_pending_updates=True
        )

    except Exception as e:
        logger.error(f"❌ Error menjalankan bot: {e}")
        print(f"❌ Error: {e}")

if __name__ == '__main__':
    main()
