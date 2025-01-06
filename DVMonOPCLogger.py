#!/usr/bin/env python3
# DVMonOPCLogger.py
#
# Python Version: 3.13
#
# Command-line utility for reading OPC DA tags from a file (.csv/.xlsx/.xls).
# This can be compiled into DVMonOPCLogger.exe using PyInstaller or a similar tool.

import sys
import argparse
import logging
import time

try:
    import OpenOPC
except ImportError:
    OpenOPC = None
    print("Note: 'OpenOPC' is not installed. Please install via 'pip install OpenOPC-Python3x' if using OPC DA.")

try:
    import polars as pl
except ImportError:
    print("Polars is not installed. Please install via 'pip install polars' to proceed.")
    sys.exit(1)

__version__ = "1.0.0"


def parse_args():
    """
    Parse command-line arguments.
    """
    parser = argparse.ArgumentParser(
        description="DVMonOPCLogger - A CLI OPC DA logging utility (CSV/XLSX/XLS support via Polars)."
    )
    parser.add_argument(
        "--filename",
        type=str,
        help="Path to the file containing columns: Tag, Timestamp, Read Value."
    )
    parser.add_argument(
        "--version",
        action="store_true",
        help="Display version information."
    )
    parser.add_argument(
        "--info",
        action="store_true",
        help="Display information about this tool."
    )
    parser.add_argument(
        "--client",
        type=str,
        help="OPC DA client name (e.g., 'Matrikon.OPC.Simulation.1')."
    )
    parser.add_argument(
        "--max_tags_per_interval",
        type=int,
        default=100,
        help="Maximum number of tags read per interval (default=100)."
    )
    parser.add_argument(
        "--interval_seconds",
        type=int,
        default=60,
        help="Interval in seconds between reading tag batches (default=60)."
    )

    return parser.parse_args()


def display_version():
    """Prints the version of this tool."""
    print(f"DVMonOPCLogger version: {__version__}")


def display_info():
    """Prints information about this tool."""
    info_text = (
        "DVMonOPCLogger - A command line utility to load a set of OPC DA tags from a file,\n"
        "connect to an OPC DA server, read their values, and log or process them as needed.\n\n"
        "Supported file format (via Polars):\n"
        "  CSV/XLSX/XLS with columns: Tag, Timestamp, Read Value.\n\n"
        "Optional arguments:\n"
        "  --client                 OPC DA client name (ProgID)\n"
        "  --max_tags_per_interval  Number of tags to read per interval (default=100)\n"
        "  --interval_seconds       Seconds between reading intervals (default=60)\n"
    )
    print(info_text)


def read_tags_file(filename):
    """
    Uses Polars to read a file (.csv/.xlsx/.xls) containing columns:
    Tag, Timestamp, Read Value.

    We only need the 'Tag' column to read from the OPC server.
    The 'Timestamp' and 'Read Value' columns are ignored in this script.
    """
    if not filename:
        logging.error("No filename specified.")
        return []

    # Determine file extension
    file_ext = filename.split('.')[-1].lower()

    try:
        if file_ext == 'csv':
            df = pl.read_csv(filename)
        elif file_ext in ('xlsx', 'xls'):
            # Polars read_excel typically supports .xlsx
            # .xls may fail depending on Polars/xlrd support.
            df = pl.read_excel(filename)
        else:
            logging.error("Unsupported file extension. Use .csv, .xlsx, or .xls.")
            sys.exit(1)
    except Exception as e:
        logging.error(f"Error reading '{filename}' with Polars: {e}")
        sys.exit(1)

    # We assume the file has columns named exactly: Tag, Timestamp, Read Value
    # If your columns differ, adjust accordingly.
    # Only read the 'Tag' column to retrieve the tag list for OPC reads.
    if 'Tag' not in df.columns:
        logging.error("The required 'Tag' column is missing from the file.")
        sys.exit(1)

    # Convert to list, dropping nulls and duplicates, and strip whitespace
    tag_list = (
        df['Tag']
        .drop_nulls()
        .unique()
        .to_list()
    )

    # Return cleaned tags
    return [str(tag).strip() for tag in tag_list if str(tag).strip()]


def connect_to_opc_server(client_name):
    """
    Connect to the specified OPC DA server using OpenOPC.

    :param client_name: The ProgID (or other identifier) for the OPC DA server.
    :return: An OpenOPC client connection, or None if not successful.
    """
    if OpenOPC is None:
        logging.error("OpenOPC is not available. Please install it and try again.")
        return None

    if not client_name:
        logging.error("No OPC DA client name provided. Use --client <client_name>.")
        return None

    try:
        opc = OpenOPC.client()
        opc.connect(client_name)
        logging.info(f"Connected to OPC DA server: {client_name}")
        return opc
    except Exception as e:
        logging.error(f"Failed to connect to OPC DA server '{client_name}'. Error: {e}")
        return None


def read_values_from_opc(opc_connection, tag_list):
    """
    Read values for each tag from an OPC DA server using OpenOPC.

    :param opc_connection: The OpenOPC client connection.
    :param tag_list: List of tag names (strings) to read.
    :return: Dictionary mapping tag -> value.
    """
    if not opc_connection:
        logging.error("OPC connection is None. Cannot read tags.")
        return {}

    try:
        read_result = opc_connection.read(tag_list)  # returns a list of tuples
        values = {}
        for item in read_result:
            tag, value, quality, timestamp = item
            values[tag] = value
        return values
    except Exception as e:
        logging.error(f"Error reading tags: {e}")
        return {}


def chunk_list(data, chunk_size):
    """Utility generator that yields chunks of size `chunk_size` from `data`."""
    for i in range(0, len(data), chunk_size):
        yield data[i : i + chunk_size]


def main():
    args = parse_args()

    if args.version:
        display_version()
        sys.exit(0)

    if args.info:
        display_info()
        sys.exit(0)

    if not args.filename:
        print("No --filename provided. Use --info for more help.")
        sys.exit(1)

    logging.basicConfig(level=logging.INFO)

    # Read the tags from the file using Polars (only the 'Tag' column)
    tags = read_tags_file(args.filename)
    if not tags:
        logging.warning("No tags found in the specified file.")
        sys.exit(1)

    # Connect to the OPC server using the provided client name
    opc_conn = connect_to_opc_server(args.client)
    if not opc_conn:
        logging.error("Failed to establish an OPC DA connection. Exiting.")
        sys.exit(1)

    # Process tags in chunks
    max_tags = args.max_tags_per_interval
    interval = args.interval_seconds
    logging.info(
        f"Processing up to {max_tags} tags every {interval} seconds until all tags are read."
    )

    for idx, chunk in enumerate(chunk_list(tags, max_tags), start=1):
        logging.info(f"Reading chunk {idx}: {len(chunk)} tags.")

        values = read_values_from_opc(opc_conn, chunk)
        for tag, value in values.items():
            logging.info(f"  {tag} = {value}")

        # If this is not the last chunk, wait before reading the next one
        if idx * max_tags < len(tags):
            logging.info(f"Waiting {interval} seconds before the next chunk...")
            time.sleep(interval)

    # Disconnect from OPC server
    if opc_conn:
        opc_conn.close()
        logging.info("Disconnected from OPC server.")


if __name__ == "__main__":
    main()
