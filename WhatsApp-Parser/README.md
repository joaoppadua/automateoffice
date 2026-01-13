# WhatsApp Chat Parser

A Python script to parse WhatsApp backup chat files (.txt) and export them to structured formats (CSV and Excel).

## Features

- Parse WhatsApp chat backup .txt files
- Extract date, time, author, and message content
- Handle multi-line messages
- Support for backup creator notation (shown as ".")
- Generate separate files for each chat participant
- Export to both CSV and Excel formats
- Display summary statistics

## Installation

### Prerequisites

- Python 3.10 or higher
- Conda (recommended) or pip

### Setup with Conda

1. Clone this repository:
```bash
git clone <repository-url>
cd WhatsApp-Parser
```

2. Create a conda environment:
```bash
conda create -n whatsapp-parser python=3.10
conda activate whatsapp-parser
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

### Setup with pip

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

By default, output files are created in the same directory as the input file, using the input filename as the prefix:

```bash
python whatsapp_parser.py <path_to_chat_file.txt>
```

For example, if your input file is `/home/user/chats/family_chat.txt`, the output files will be created as:
- `/home/user/chats/family_chat.csv`
- `/home/user/chats/family_chat.xlsx`

### Custom Output Prefix

You can specify a custom output location and prefix using the `-o` flag:

```bash
python whatsapp_parser.py <path_to_chat_file.txt> -o my_custom_prefix
```

### Example

```bash
# Use default output (same directory and name as input)
python whatsapp_parser.py /home/user/chats/conversation.txt

# Specify custom output prefix
python whatsapp_parser.py chat_backup.txt -o family_chat
```

## Input Format

The script expects WhatsApp chat backup files in the standard format:

```
[DD/MM/YYYY, HH:MM:SS] Author Name: Message content
[15/03/2024, 14:30:45] John Smith: Are we still meeting at the cafe?
[15/03/2024, 14:31:02] .: Yes, see you at 3pm
```

**Note:** When the message is from the person who created the backup, WhatsApp shows a dot (`.`) instead of the author's name.

## Output Files

The script generates the following files:

### Main Chat Export
- `<output_prefix>.csv` - All messages in CSV format
- `<output_prefix>.xlsx` - All messages in Excel format

### Per-Author Exports
- `<output_prefix>_<Author1>.csv` - Messages from Author 1 (CSV)
- `<output_prefix>_<Author1>.xlsx` - Messages from Author 1 (Excel)
- `<output_prefix>_<Author2>.csv` - Messages from Author 2 (CSV)
- `<output_prefix>_<Author2>.xlsx` - Messages from Author 2 (Excel)

Each file contains the following columns:
- **Date**: Message date (DD/MM/YYYY)
- **Time**: Message time (HH:MM:SS)
- **Author**: Name of the message sender
- **Message**: Content of the message

## Command-Line Options

```
usage: whatsapp_parser.py [-h] [-o OUTPUT] input_file

positional arguments:
  input_file            Path to WhatsApp chat backup .txt file

optional arguments:
  -h, --help            Show help message and exit
  -o OUTPUT, --output OUTPUT
                        Output file prefix (default: same directory and name as input file)
```

## Example Output

```
Parsing WhatsApp chat file: chat.txt

Total messages parsed: 1523

Authors found: John Smith, .

Exporting main chat data...
✓ Exported to CSV: chat.csv
✓ Exported to Excel: chat.xlsx

Exporting individual author data...

Author: John Smith (856 messages)
✓ Exported to CSV: chat_John_Smith.csv
✓ Exported to Excel: chat_John_Smith.xlsx

Author: . (667 messages)
✓ Exported to CSV: chat__.csv
✓ Exported to Excel: chat__.xlsx

✓ All exports completed successfully!

==================================================
SUMMARY
==================================================
Total messages: 1523

Messages per author:
John Smith     856
.              667
```

## Troubleshooting

### File not found error
Ensure the path to your WhatsApp chat file is correct. Use absolute paths if needed:
```bash
python whatsapp_parser.py /full/path/to/chat.txt
```

### No messages parsed
Check that your file follows the WhatsApp backup format. The script looks for lines starting with `[DD/MM/YYYY, HH:MM:SS]`.

### Encoding issues
The script uses UTF-8 encoding. If you encounter encoding errors, ensure your chat file is saved in UTF-8 format.

## Dependencies

- pandas >= 2.0.0
- openpyxl >= 3.1.0

## License

This project is open source and available for personal and educational use.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.
