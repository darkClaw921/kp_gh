import argparse
import base64
import csv
import json
import os
import re
from html import unescape
from pathlib import Path
from urllib.parse import urljoin, urlparse
import requests


def _extract_first(pattern: str, text: str, flags: int = 0) -> str | None:
    match = re.search(pattern, text, flags)
    return match.group(1).strip() if match else None


def _download_image_to_base64(image_url: str, base_url: str = "") -> dict | None:
    """Загружает изображение и возвращает его в формате base64 с именем файла"""
    try:
        # Если URL относительный, делаем его абсолютным
        if not image_url.startswith(('http://', 'https://')):
            if base_url:
                image_url = urljoin(base_url, image_url)
            else:
                return None
        
        response = requests.get(image_url, timeout=10)
        response.raise_for_status()
        
        # Получаем имя файла из URL
        parsed_url = urlparse(image_url)
        filename = os.path.basename(parsed_url.path)
        if not filename or '.' not in filename:
            filename = f"image_{hash(image_url) % 10000}.jpg"
        
        # Конвертируем в base64
        encoded_data = base64.b64encode(response.content).decode('utf-8')
        
        return {
            'fileData': [filename, encoded_data]
        }
    except Exception:
        return None


def parse_catalog_items(html_text: str, base_url: str = "", download_images: bool = False) -> list[dict]:
    items: list[dict] = []

    # Разбиваем по карточкам. Первый элемент до первой карточки — пропускаем
    parts = html_text.split('<div class="catalog-item"')
    if len(parts) <= 1:
        return items

    for part in parts[1:]:
        chunk = '<div class="catalog-item"' + part

        # Фото (img data-src) — приоритетно берем из <img class="lazy" data-src="...">
        img_url = _extract_first(r'<img[^>]*class="lazy"[^>]*data-src="([^"]+)"', chunk, re.DOTALL)
        if not img_url:
            # запасной вариант — первый URL из data-srcset у <source>
            srcset = _extract_first(r'<source[^>]*data-srcset="([^"]+)"', chunk, re.DOTALL)
            if srcset:
                img_url = srcset.split(',')[0].split()[0].strip()

        # Ссылка на товар — берем первую ссылку, ведущую в /catalog/
        link = _extract_first(r'<a\s+href="(/catalog/[^"]+)"', chunk)

        # Название — внутри div.catalog-item-name
        name = _extract_first(r'<div\s+class="catalog-item-name">\s*(.*?)\s*</div>', chunk, re.DOTALL)

        # Цена — внутри span.catalog-item-price
        price = _extract_first(r'<span\s+class="catalog-item-price">\s*(.*?)\s*</span>', chunk, re.DOTALL)

        if name or price or link or img_url:
            item_data = {
                "name": unescape(name) if name else None,
                "price": unescape(price) if price else None,
                "link": unescape(link) if link else None,
                "image": unescape(img_url) if img_url else None,
            }
            
            # Если нужно загружать изображения в base64
            if download_images and img_url:
                image_base64 = _download_image_to_base64(img_url, base_url)
                if image_base64:
                    item_data["image_base64"] = image_base64
            
            items.append(item_data)

    return items


def _write_csv(path: Path, rows: list[dict]) -> None:
    cols = ["name", "price", "link", "image", "image_base64"]
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=cols)
        writer.writeheader()
        for row in rows:
            # Для CSV base64 данные не выводим (слишком большие)
            csv_row = {c: (row.get(c) or "") for c in cols if c != "image_base64"}
            writer.writerow(csv_row)


def main() -> None:
    parser = argparse.ArgumentParser(description="Парсинг карточек товара из HTML каталога")
    parser.add_argument("input", nargs="?", default=str(Path(__file__).with_name("page_2.html")), help="Путь к HTML-файлу (по умолчанию parser_site/page_1.html)")
    parser.add_argument("--out", dest="out", default=None, help="Путь для сохранения результата (JSON/CSV определяется по расширению)")
    parser.add_argument("--format", dest="fmt", choices=["json", "csv"], default="json", help="Формат вывода, если --out не задан или без расширения")
    parser.add_argument("--base-url", dest="base_url", default="", help="Базовый URL для загрузки изображений (если они относительные)")
    parser.add_argument("--download-images", action="store_true", help="Загружать изображения и конвертировать в base64")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"Файл не найден: {input_path}")

    html_text = input_path.read_text(encoding="utf-8", errors="ignore")
    items = parse_catalog_items(html_text, args.base_url, args.download_images)

    if args.out:
        out_path = Path(args.out)
        suffix = out_path.suffix.lower()
        if suffix == ".csv" or (suffix == "" and args.fmt == "csv"):
            _write_csv(out_path if suffix else out_path.with_suffix(".csv"), items)
        else:
            out_json = out_path if suffix == ".json" or suffix == "" else out_path.with_suffix(".json")
            out_json.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        if args.fmt == "csv":
            # Печать CSV в stdout
            cols = ["name", "price", "link", "image"]
            writer = csv.DictWriter(
                f=__import__("sys").stdout,
                fieldnames=cols,
            )
            writer.writeheader()
            for row in items:
                writer.writerow({c: (row.get(c) or "") for c in cols})
        else:
            print(json.dumps(items, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    # С загрузкой изображений в base64
    # python3 parser_site/parse_page.py --download-images --base-url "https://gusevskoe-steklo.ru" --out items.json
    # или в ручную 
    
    main()
    


