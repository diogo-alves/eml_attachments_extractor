import mimetypes
import re
from argparse import ArgumentParser, ArgumentTypeError
from email import message_from_file, policy
from pathlib import Path
from typing import List


def generate_unique_filename(filepath: Path) -> Path:
    # Generate a unique filename by adding an incrementing number
    if not filepath.exists():
        return filepath

    base = filepath.stem
    suffix = filepath.suffix
    counter = 1

    while True:
        new_filepath = filepath.with_name(f"{base}_{counter}{suffix}")
        if not new_filepath.exists():
            return new_filepath
        counter += 1


def extract_attachments(file: Path, destination: Path) -> None:
    print(f'PROCESSING FILE "{file}"')
    try:
        with file.open(encoding='utf-8', errors='replace') as f:
            email_message = message_from_file(f, policy=policy.default)
            email_subject = email_message.get('Subject', 'No Subject')
            basepath = destination / sanitize_foldername(email_subject)
            duplicates_path = destination / "duplicates"
            # ignore inline attachments
            attachments = [item for item in email_message.iter_attachments() if item.is_attachment()]  # type: ignore
            if not attachments:
                print('>> No attachments found.')
                return

            # Create base directory before processing attachments
            basepath.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                filename = get_safe_filename(attachment)
                if not filename:
                    print('>> Attachment found: None')
                    continue

                print(f'>> Attachment found: {filename}')
                filepath = basepath / filename
                payload = attachment.get_payload(decode=True)

                if filepath.exists():
                    print('>> Moving to duplicates folder...')
                    duplicates_path.mkdir(exist_ok=True)
                    duplicate_filepath = generate_unique_filename(duplicates_path / filename)
                    save_attachment(duplicate_filepath, payload)
                else:
                    save_attachment(filepath, payload)
    except UnicodeDecodeError:
        # If UTF-8 fails, try with binary mode
        with file.open('rb') as f:
            email_message = message_from_file(f, policy=policy.default)
            email_subject = email_message.get('Subject', 'No Subject')
            basepath = destination / sanitize_foldername(email_subject)
            duplicates_path = destination / "duplicates"

            attachments = [item for item in email_message.iter_attachments() if item.is_attachment()]
            if not attachments:
                print('>> No attachments found.')
                return

            # Create base directory before processing attachments
            basepath.mkdir(parents=True, exist_ok=True)

            for attachment in attachments:
                filename = get_safe_filename(attachment)
                if not filename:
                    print('>> Attachment found: None')
                    continue

                print(f'>> Attachment found: {filename}')
                filepath = basepath / filename
                payload = attachment.get_payload(decode=True)

                if filepath.exists():
                    print('>> Moving to duplicates folder...')
                    duplicates_path.mkdir(exist_ok=True)
                    duplicate_filepath = generate_unique_filename(duplicates_path / filename)
                    save_attachment(duplicate_filepath, payload)
                else:
                    save_attachment(filepath, payload)


def sanitize_foldername(name: str) -> str:
    if not name:
        return "unnamed"
    # Remove or replace characters that are problematic in file paths
    illegal_chars = r'[/\\|\[\]\{\}:<>+=;,?!*"~#$%&@\']'
    sanitized = re.sub(illegal_chars, '_', name)
    # Remove trailing spaces and periods which can cause issues on Windows
    sanitized = sanitized.strip('. ')
    # Replace multiple spaces with single space
    sanitized = re.sub(r'\s+', ' ', sanitized)
    # Limit the length to avoid path length issues
    return sanitized[:100]


def save_attachment(file: Path, payload: bytes) -> None:
    try:
        # Ensure parent directory exists
        file.parent.mkdir(parents=True, exist_ok=True)
        with file.open('wb') as f:
            print(f'>> Saving attachment to "{file}"')
            f.write(payload)
    except Exception as e:
        print(f'>> Error saving attachment: {e}')


def get_safe_filename(attachment) -> str:
    try:
        filename = attachment.get_filename()
        if not filename:
            return None

        # Handle RFC 2231 encoded filenames
        filename = filename.replace('"', '').replace('\t', '')
        filename = re.sub(r';\s*filename\*\d*=', '', filename)

        # Basic sanitization
        filename = sanitize_foldername(filename)

        # Ensure filename has an extension
        if not Path(filename).suffix and attachment.get_content_type():
            ext = mimetypes.guess_extension(attachment.get_content_type())
            if ext:
                filename = f"{filename}{ext}"

        return filename
    except Exception as e:
        print(f"Error processing filename: {e}")
        return None


def get_eml_files_from(path: Path, recursively: bool = False) -> List[Path]:
    if recursively:
        return list(path.rglob('*.eml'))
    return list(path.glob('*.eml'))


def check_file(arg_value: str) -> Path:
    file = Path(arg_value)
    if file.is_file() and file.suffix == '.eml':
        return file
    raise ArgumentTypeError(f'"{file}" is not a valid EML file.')


def check_path(arg_value: str) -> Path:
    path = Path(arg_value)
    if path.is_dir():
        return path
    raise ArgumentTypeError(f'"{path}" is not a valid directory.')


def main():
    parser = ArgumentParser(
        usage='%(prog)s [OPTIONS]',
        description='Extracts attachments from .eml files'
    )
    # force the use of --source or --files, not both
    source_group = parser.add_mutually_exclusive_group()
    source_group.add_argument(
        '-s',
        '--source',
        type=check_path,
        default=Path.cwd(),
        metavar='PATH',
        help='the directory containing the .eml files to extract attachments (default: current working directory)'
    )
    parser.add_argument(
        '-r',
        '--recursive',
        action='store_true',
        help='allow recursive search for .eml files under SOURCE directory'
    )
    source_group.add_argument(
        '-f',
        '--files',
        nargs='+',
        type=check_file,
        metavar='FILE',
        help='specify a .eml file or a list of .eml files to extract attachments'
    )
    parser.add_argument(
        '-d',
        '--destination',
        type=check_path,
        default=Path.cwd(),
        metavar='PATH',
        help='the directory to extract attachments into (default: current working directory)'
    )
    args = parser.parse_args()

    eml_files = args.files or get_eml_files_from(args.source, args.recursive)
    if not eml_files:
        print(f'No EML files found!')

    for file in eml_files:
        extract_attachments(file, destination=args.destination)
    print('Done.')


if __name__ == '__main__':
    main()
