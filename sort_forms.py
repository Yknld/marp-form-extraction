"""Copy forms into format folders. No OCR, downloads, or source modifications."""
import argparse
import csv
from collections import Counter
from pathlib import Path
import shutil

FOLDERS = ('Word documents', 'Text PDFs', 'Needs OCR', 'Other or unreadable files')
WORD = {'.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm', '.rtf', '.odt'}
IMAGES = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.heic', '.webp'}


def classify(path):
    suffix = path.suffix.lower()
    if suffix in WORD:
        return FOLDERS[0], 'Word/document format; contents not inspected'
    if suffix in IMAGES:
        return FOLDERS[2], 'Image file'
    if suffix != '.pdf':
        return FOLDERS[3], 'Other file format'
    import pdfplumber
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return FOLDERS[3], 'PDF has no pages'
            readable = False
            for number, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ''
                letters = sum(c.isalnum() for c in text)
                readable |= letters >= 20
                # Short headers/page numbers do not make a scanned page readable.
                area = max(1, float(page.width * page.height))
                large_image = any(
                    max(0, im['x1'] - im['x0']) * max(0, im['bottom'] - im['top']) / area > .25
                    for im in page.images)
                if (large_image and letters < 150) or (page.images and letters < 20):
                    return FOLDERS[2], f'Page {number} appears scanned; entire PDF needs review'
                if letters < 20 and page.curves:
                    return FOLDERS[2], f'Page {number} has graphics but little readable text'
            if readable:
                return FOLDERS[1], 'Readable text detected; extraction quality not verified'
            return FOLDERS[2], 'Little or no readable text; may be blank or need OCR'
    except Exception as exc:
        return FOLDERS[3], f'PDF could not be read ({type(exc).__name__})'


def run(source, output):
    source, output = Path(source).resolve(), Path(output).resolve()
    if not source.is_dir():
        raise ValueError('Input must be an existing folder')
    if source == output or source in output.parents or output in source.parents:
        raise ValueError('Use separate input and output folders, neither inside the other')
    # Check the one dependency before copying anything. Never install it automatically.
    import pdfplumber  # noqa: F401
    files = sorted(p for p in source.rglob('*') if p.is_file() and not p.is_symlink())
    for folder in FOLDERS:
        (output / folder).mkdir(parents=True, exist_ok=True)
    counts = Counter()
    with (output / 'sorting_report.csv').open('w', newline='', encoding='utf-8-sig') as report:
        writer = csv.writer(report)
        writer.writerow(['Source file', 'Folder', 'Result', 'Reason'])
        for i, path in enumerate(files, 1):
            relative = path.relative_to(source)
            if any(part.startswith(('.', '~$')) or part == '__MACOSX' for part in relative.parts):
                writer.writerow([str(relative), '', 'Skipped', 'Hidden or temporary file'])
                report.flush()
                continue
            print(f'[{i}/{len(files)}] Checking {relative}...', flush=True)
            folder, reason = classify(path)
            destination = output / folder / relative
            try:
                destination.parent.mkdir(parents=True, exist_ok=True)
                # Exclusive creation protects existing copies and manual edits.
                with destination.open('xb') as target:
                    try:
                        with path.open('rb') as original:
                            shutil.copyfileobj(original, target)
                    except BaseException:
                        target.close()
                        destination.unlink()
                        raise
                result = 'Copied'
            except FileExistsError:
                result = 'Skipped existing copy'
            except OSError as exc:
                result = 'Copy failed'
                reason += f'; {type(exc).__name__}'
            counts[result] += 1
            writer.writerow([str(relative), folder, result, reason])
            report.flush()
            print(f'  {folder}: {result}', flush=True)
    print(f'Saved to {output}\n' + '; '.join(f'{k}: {v}' for k, v in counts.items()), flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    try:
        run(args.input, args.output)
    except ImportError:
        parser.exit(2, 'pdfplumber is unavailable in this Python environment. Use the environment running your extractor; nothing was installed.\n')
    except (ValueError, OSError) as exc:
        parser.exit(2, f'Unable to sort: {exc}\n')
    except KeyboardInterrupt:
        parser.exit(130, 'Stopped. Completed copies remain; originals are unchanged.\n')


if __name__ == '__main__':
    main()
