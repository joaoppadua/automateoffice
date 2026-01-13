#!/usr/bin/env python3
"""
WhatsApp Chat Parser
Parses WhatsApp backup .txt files and exports to CSV and Excel formats.
"""

import re
import pandas as pd
from datetime import datetime
import argparse
import sys
from pathlib import Path


def parse_whatsapp_chat(file_path):
    """
    Parse WhatsApp chat backup file.

    Args:
        file_path: Path to the WhatsApp .txt backup file

    Returns:
        pandas.DataFrame: DataFrame with columns Date, Time, Author, Message
    """
    # Pattern to match WhatsApp message format: [DD/MM/YYYY, HH:MM:SS] Author: Message
    message_pattern = re.compile(r'^\[(\d{2}/\d{2}/\d{4}),\s*(\d{2}:\d{2}:\d{2})\]\s*([^:]+):\s*(.*)$')

    messages = []
    current_message = None

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.rstrip('\n')
                match = message_pattern.match(line)

                if match:
                    # Save previous message if exists
                    if current_message:
                        messages.append(current_message)

                    # Start new message
                    date, time, author, content = match.groups()
                    current_message = {
                        'Date': date,
                        'Time': time,
                        'Author': author.strip(),
                        'Message': content
                    }
                else:
                    # Continuation of previous message (multi-line message)
                    if current_message:
                        current_message['Message'] += '\n' + line

            # Don't forget the last message
            if current_message:
                messages.append(current_message)

    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

    if not messages:
        print("Warning: No messages found in the file.")
        return pd.DataFrame(columns=['Date', 'Time', 'Author', 'Message'])

    df = pd.DataFrame(messages)
    return df


def export_data(df, output_prefix):
    """
    Export DataFrame to CSV and Excel formats.

    Args:
        df: pandas DataFrame to export
        output_prefix: Prefix for output files (without extension)
    """
    # Export to CSV
    csv_file = f"{output_prefix}.csv"
    df.to_csv(csv_file, index=False, encoding='utf-8')
    print(f"✓ Exported to CSV: {csv_file}")

    # Export to Excel
    excel_file = f"{output_prefix}.xlsx"
    df.to_excel(excel_file, index=False, engine='openpyxl')
    print(f"✓ Exported to Excel: {excel_file}")


def main():
    parser = argparse.ArgumentParser(
        description='Parse WhatsApp chat backup files and export to CSV/Excel'
    )
    parser.add_argument(
        'input_file',
        help='Path to WhatsApp chat backup .txt file'
    )
    parser.add_argument(
        '-o', '--output',
        default='whatsapp_chat',
        help='Output file prefix (default: whatsapp_chat)'
    )

    args = parser.parse_args()

    print(f"Parsing WhatsApp chat file: {args.input_file}")

    # Parse the chat file
    df_all = parse_whatsapp_chat(args.input_file)

    if df_all.empty:
        print("No data to export.")
        sys.exit(1)

    print(f"\nTotal messages parsed: {len(df_all)}")

    # Get unique authors
    authors = df_all['Author'].unique()
    print(f"Authors found: {', '.join(authors)}")

    # Export main DataFrame
    print("\nExporting main chat data...")
    export_data(df_all, args.output)

    # Export separate DataFrames for each author
    if len(authors) > 0:
        print("\nExporting individual author data...")
        for author in authors:
            df_author = df_all[df_all['Author'] == author].copy()
            # Sanitize author name for filename
            safe_author = re.sub(r'[^\w\s-]', '_', author).strip()
            output_prefix = f"{args.output}_{safe_author}"
            print(f"\nAuthor: {author} ({len(df_author)} messages)")
            export_data(df_author, output_prefix)

    print("\n✓ All exports completed successfully!")

    # Display summary statistics
    print("\n" + "="*50)
    print("SUMMARY")
    print("="*50)
    print(f"Total messages: {len(df_all)}")
    print("\nMessages per author:")
    print(df_all['Author'].value_counts().to_string())


if __name__ == "__main__":
    main()
