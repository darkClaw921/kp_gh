import pdfkit
from pypdf import PdfReader, PdfWriter
import os
import re
import base64
from io import BytesIO
from PIL import Image

def insert_html_page_to_pdf(
    html_path='индивидуальный_расчет.html',
    base_pdf_path='кп 3 вариант.pdf',
    output_pdf_path='кп 3 вариант_с_вставкой.pdf',
    insert_after_page=2
):
    # 1. Подготовим временную копию HTML: исправим MIME для data URI (image/jpg -> image/jpeg)
    #    Это важно, т.к. wkhtmltopdf может не отрисовывать image/jpg.
    temp_html = 'temp_centered.html'
    def _transcode_data_uris_to_png(html: str) -> str:
        # 1) выправляем некорректный MIME
        fixed = html.replace('image/jpg', 'image/jpeg')

        # 2) находим все data:image/*;base64, ...
        pattern = re.compile(r"data:image/(png|jpeg|jpg|webp);base64,([A-Za-z0-9+/=]+)")

        def _replace(match: re.Match) -> str:
            mime = match.group(1)
            b64 = match.group(2)
            try:
                raw = base64.b64decode(b64)
                with Image.open(BytesIO(raw)) as img:
                    # конвертируем в RGB для совместимости (убираем альфу)
                    if img.mode not in ("RGB", "L"):
                        img = img.convert("RGB")
                    out = BytesIO()
                    img.save(out, format="PNG")
                    out_b64 = base64.b64encode(out.getvalue()).decode('ascii')
                    return f"data:image/png;base64,{out_b64}"
            except Exception:
                # если не удалось — возвращаем как есть
                return match.group(0)

        return pattern.sub(_replace, fixed)

    try:
        with open(html_path, 'r', encoding='utf-8') as f_in:
            html_content = f_in.read()
        processed_html = _transcode_data_uris_to_png(html_content)
        with open(temp_html, 'w', encoding='utf-8') as f_out:
            f_out.write(processed_html)
    except Exception:
        # Если что-то пошло не так, используем оригинальный HTML
        temp_html = html_path

    # 2. Конвертируем HTML в PDF с дефолтными настройками (пусть wkhtmltopdf сам подберет размер)
    temp_pdf = 'temp_insert_page.pdf'
    options = {
        'orientation': 'Landscape',
        'margin-top': '70mm',
        'margin-bottom': '0mm',
        'margin-left': '0mm',
        'margin-right': '50mm',
        'enable-local-file-access': None,
        'encoding': 'utf-8',
    }
    pdfkit.from_file(temp_html, temp_pdf, options=options)
    # 1/0
    # 3. Открываем исходный PDF и PDF с html-страницей
    base_reader = PdfReader(base_pdf_path)
    insert_reader = PdfReader(temp_pdf)
    writer = PdfWriter()

    # 4. Копируем первые insert_after_page+1 страниц
    for i in range(insert_after_page+1):
        writer.add_page(base_reader.pages[i])

    # 5. Копируем настройки предыдущей страницы
    prev_page = base_reader.pages[insert_after_page]
    mediabox = prev_page.mediabox
    rotation = prev_page.get("/Rotate", 0)

    # 6. Вставляем страницу из html и применяем настройки
    insert_page = insert_reader.pages[-1]
    insert_page.mediabox.lower_left = mediabox.lower_left
    insert_page.mediabox.lower_right = mediabox.lower_right
    insert_page.mediabox.upper_left = mediabox.upper_left
    insert_page.mediabox.upper_right = mediabox.upper_right
    if rotation:
        insert_page.rotate(rotation)
    writer.add_page(insert_page)

    # 7. Копируем оставшиеся страницы
    for i in range(insert_after_page+1, len(base_reader.pages)):
        writer.add_page(base_reader.pages[i])

    # 8. Сохраняем результат
    with open(output_pdf_path, 'wb') as f:
        writer.write(f)

    # 9. Удаляем временные файлы
    # os.remove(temp_pdf)
    # if temp_html != html_path:
    #     os.remove(temp_html)

if __name__ == "__main__":
    insert_html_page_to_pdf() 