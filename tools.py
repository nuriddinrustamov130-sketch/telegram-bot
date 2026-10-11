import subprocess
import json
import os
import glob
import urllib.request
import urllib.parse
from pathlib import Path
from duckduckgo_search import DDGS
from docx import Document
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pypdf
import yt_dlp

def create_and_send_word_document(filename: str, title: str, content: str) -> str:
    """Word (.docx) faylini yaratish."""
    try:
        if not filename.endswith(".docx"):
            filename += ".docx"
        file_path = Path(filename).resolve()
        doc = Document()
        if title:
            doc.add_heading(title, level=0)
        
        for para in content.split("\n\n"):
            p_text = para.strip()
            if p_text:
                doc.add_paragraph(p_text)
        
        doc.save(file_path)
        return f"FILE_READY:{str(file_path)}"
    except Exception as e:
        return f"Word fayl yaratishda xatolik: {str(e)}"

def create_and_send_excel_sheet(filename: str, sheet_title: str, headers: list, rows: list) -> str:
    """Professional Excel (.xlsx) jadval yaratish."""
    try:
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
        file_path = Path(filename).resolve()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title or "Jadval"

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for row_data in rows:
            ws.append(row_data)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
            for cell in row:
                cell.border = border

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        wb.save(file_path)
        return f"FILE_READY:{str(file_path)}"
    except Exception as e:
        return f"Excel yaratishda xatolik: {str(e)}"

def convert_pdf_to_word(pdf_path: str, output_docx_name: str = None) -> str:
    """PDF faylni Word (.docx) ga aylantirish."""
    try:
        p = Path(pdf_path).resolve()
        if not p.exists():
            return f"Xatolik: '{pdf_path}' topilmadi."

        reader = pypdf.PdfReader(str(p))
        doc = Document()
        doc.add_heading(f"PDF dan olingan matn: {p.name}", level=0)

        for idx, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                doc.add_heading(f"Sahifa {idx + 1}", level=2)
                for line in text.split("\n"):
                    if line.strip():
                        doc.add_paragraph(line.strip())

        out_name = output_docx_name or f"{p.stem}_konvertatsiya.docx"
        if not out_name.endswith(".docx"): out_name += ".docx"
        out_path = Path(out_name).resolve()
        doc.save(out_path)
        return f"FILE_READY:{str(out_path)}"
    except Exception as e:
        return f"PDF ni Wordga o'tkazishda xatolik: {str(e)}"

def download_instagram_or_video(url: str) -> str:
    """Instagram (Reels, Post) yoki video havolasidan videoni yuklash."""
    try:
        out_dir = Path("downloads").resolve()
        out_dir.mkdir(exist_ok=True)
        out_template = str(out_dir / "%(id)s.%(ext)s")

        ydl_opts = {
            'outtmpl': out_template,
            'format': 'best[ext=mp4]/best',
            'quiet': True,
            'no_warnings': True,
            'max_filesize': 49 * 1024 * 1024
        }

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            filename = ydl.prepare_filename(info)
            if not Path(filename).exists():
                matches = list(out_dir.glob(f"{info.get('id', '')}*"))
                if matches: filename = str(matches[0])
            return f"VIDEO_READY:{filename}"
    except Exception as e:
        return f"Video yuklashda xatolik: {str(e)}"

def search_and_download_music(query: str) -> str:
    """Musiqa qidirish va MP3/Audio qilib to'g'ridan-to'g'ri Telegramga yuklab berish."""
    out_dir = Path("downloads").resolve()
    out_dir.mkdir(exist_ok=True)

    # 1. iTunes rasmiy qidiruvi orqali audio topish (juda tez va toza audio)
    try:
        encoded = urllib.parse.quote(query)
        api_url = f"https://itunes.apple.com/search?term={encoded}&media=music&limit=1"
        req = urllib.request.Request(api_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            results = data.get("results", [])
            if results and results[0].get("previewUrl"):
                track = results[0]
                track_name = track.get("trackName", "music")
                artist = track.get("artistName", "artist")
                safe_name = f"{artist} - {track_name}".replace("/", "_").replace("\\", "_")
                dest_file = out_dir / f"{safe_name}.m4a"
                
                # Audio faylni yuklab olish
                urllib.request.urlretrieve(track["previewUrl"], str(dest_file))
                return f"AUDIO_READY:{str(dest_file)}"
    except Exception as e:
        pass

    # 2. yt-dlp orqali internetdan yuklash
    try:
        out_template = str(out_dir / "%(title)s.%(ext)s")
        ydl_opts = {
            'outtmpl': out_template,
            'format': 'bestaudio/best',
            'default_search': 'ytsearch1:',
            'quiet': True,
            'no_warnings': True,
            'max_filesize': 49 * 1024 * 1024
        }
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(query, download=True)
            entries = info.get('entries', [info])
            first = entries[0] if entries else info
            filename = ydl.prepare_filename(first)
            if not Path(filename).exists():
                matches = list(out_dir.glob("*"))
                if matches: filename = str(max(matches, key=os.path.getmtime))
            return f"AUDIO_READY:{filename}"
    except Exception as e:
        return f"Musiqa qidirish xatoligi: {str(e)}"

def send_file_to_telegram(path: str) -> str:
    """Faylni Telegramga yuborish."""
    p = Path(path).resolve()
    if p.exists():
        return f"FILE_READY:{str(p)}"
    return f"Fayl topilmadi: {path}"

def web_search(query: str, max_results: int = 5) -> str:
    """Internetdan ma'lumot qidirish."""
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=max_results))
        if not results:
            return f"'{query}' bo'yicha natija topilmadi."
        items = []
        for i, r in enumerate(results, 1):
            items.append(f"{i}. {r.get('title')}\n{r.get('body')}\nLink: {r.get('href')}")
        return "\n\n".join(items)
    except Exception as e:
        return f"Qidiruv xatoligi: {str(e)}"

def run_command(command: str, cwd: str = ".") -> str:
    """Terminal buyrug'ini bajarish."""
    try:
        proc = subprocess.run(
            command,
            shell=True,
            cwd=cwd,
            capture_output=True,
            text=True,
            timeout=60,
            encoding="utf-8",
            errors="replace"
        )
        out = proc.stdout.strip()
        err = proc.stderr.strip()
        res = []
        if out: res.append(f"Chiqish:\n{out}")
        if err: res.append(f"Xatolik:\n{err}")
        return "\n".join(res) if res else "Buyruq bajarildi."
    except Exception as e:
        return f"Buyruq xatoligi: {str(e)}"

BOT_TOOLS = [
    create_and_send_word_document,
    create_and_send_excel_sheet,
    convert_pdf_to_word,
    download_instagram_or_video,
    search_and_download_music,
    send_file_to_telegram,
    web_search,
    run_command
]
