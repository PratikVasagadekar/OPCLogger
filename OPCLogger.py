import click
import pandas as pd
import OpenOPC
import time
import logging
import sys
import gc
from datetime import datetime

VERSION = "1.1.1"

INFO_MESSAGE = """
This tool was designed to run an OPC Reader utility to get values from DeltaV OPC Server in burst modes.
It then saves the data into the same CSV/XLSX file it was provided with. The input file must contain a column named "Tag".

## Argument Requirements

- --tagfile: Provide a CSV or XLSX file with a column called "Tag" and add the tags that need to be read from DeltaV / OPC Server.

- --servername: In case you wish to read the values from a different OPC server, you can provide the OPC server string here.

- --maxtagsperinterval: Provide a number that will be used to read the tags in burst mode.

- --intervalseconds: Provide a number in seconds that will be used to wait until the next burst runs.

- --disconnect_wait_time : Wait time for the OPC Conn to get disconnected.

"""

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

def read_tags(filepath):
    """
    Reads an Excel or CSV file containing a 'Tag' column.
    Returns a list of tag strings.
    """
    try:
        if filepath.endswith('.xlsx'):
            df = pd.read_excel(filepath, engine='openpyxl')
        elif filepath.endswith('.csv'):
            df = pd.read_csv(filepath)
        else:
            raise ValueError("Unsupported file format. Please provide a .xlsx or .csv file.")
    except Exception as e:
        raise ValueError(f"Error reading the tag file: {e}")

    if 'Tag' not in df.columns:
        raise ValueError("The input file must contain a 'Tag' column.")

    return df['Tag'].tolist()

def parse_timestamp(raw_ts):
    """
    Convert the OPC timestamp string (e.g. '06/24/07 17:44:43') 
    to the format "DD-MM-YYYY HH:MM:SS AM/PM".
    
    If the timestamp from OPC has a different format, adjust accordingly.
    """
    try:
        # The original OPC timestamp is typically in the format: mm/dd/yy HH:MM:SS
        dt = datetime.strptime(raw_ts, '%m/%d/%y %H:%M:%S')
        return dt.strftime('%d-%m-%Y %I:%M:%S %p')
    except ValueError:
        # If parsing fails or format is different, return as-is or handle differently.
        return raw_ts

def write_values(filepath, tag_values_map):
    """
    Writes the provided tag information into the same file as columns:
    'Value', 'Status', 'Timestamp'.

    :param filepath: The original file path (CSV or XLSX).
    :param tag_values_map: A dict mapping tag -> (value, status, timestamp_in_requested_format).
    """
    try:
        if filepath.endswith('.xlsx'):
            # Read existing Excel file
            df = pd.read_excel(filepath, engine='openpyxl')
        elif filepath.endswith('.csv'):
            # Read existing CSV file
            df = pd.read_csv(filepath)
        else:
            raise ValueError("Unsupported file format. Please provide a .xlsx or .csv file.")

        # Create the columns if they don't exist yet
        if 'Value' not in df.columns:
            df['Value'] = None
        if 'Status' not in df.columns:
            df['Status'] = None
        if 'Timestamp' not in df.columns:
            df['Timestamp'] = None

        # Update each row by matching Tag to the dictionary
        for index, row in df.iterrows():
            tag = row['Tag']
            if tag in tag_values_map:
                val, status, timestamp_str = tag_values_map[tag]
                df.at[index, 'Value'] = val
                df.at[index, 'Status'] = status
                df.at[index, 'Timestamp'] = timestamp_str

        # Write the updated DataFrame back to the file
        if filepath.endswith('.xlsx'):
            df.to_excel(filepath, index=False, engine='openpyxl')
        else:
            df.to_csv(filepath, index=False)

    except Exception as e:
        raise ValueError(f"Error writing to the tag file: {e}")

class OPCHandler:
    def __init__(self, servername, maxtags, interval, tags, filepath, logger, disconnect_wait_time):
        self.servername = servername
        self.maxtags = maxtags
        self.interval = interval
        self.tags = tags
        self.filepath = filepath
        self.logger = logger
        self.disconnect_wait_time = disconnect_wait_time
        self.opc = None  # Initialize OPC client as None

    def connect(self):
        try:
            # Reinitialize OPC client for each batch
            self.opc = OpenOPC.client()
            self.opc.connect(self.servername)
            self.logger.info(f"Connected to OPC server: {self.servername}")
        except Exception as e:
            self.logger.error(f"Failed to connect to OPC server: {e}")
            sys.exit(1)

    def close_connection(self):
        try:
            if self.opc is not None:
                self.opc.close()
                self.logger.info("Closed OPC connection.")

            # Force OPC client deletion and garbage collection
            del self.opc
            self.opc = None
            gc.collect()
            self.logger.info(f"Waiting {self.disconnect_wait_time} seconds before reconnecting...")
            time.sleep(self.disconnect_wait_time)  # Allow DeltaV to release the tag allocation

        except Exception as e:
            self.logger.error(f"Error closing OPC connection: {e}")

    def run(self):
        """
        Reads tags in batches while respecting the maxtags limit and interval.
        Each batch establishes a fresh OPC connection, reads the tags, writes data, 
        closes the connection, and waits for a specified time before the next batch.
        """
        total_tags = len(self.tags)
        batches = [self.tags[i:i+self.maxtags] for i in range(0, total_tags, self.maxtags)]
        total_batches = len(batches)
        current_batch = 1

        while current_batch <= total_batches:
            self.logger.info(f"Processing batch {current_batch} of {total_batches}")
            batch = batches[current_batch - 1]

            self.connect()
            try:
                tag_values_map = {}

                # Read the current batch of tags
                values = self.opc.read(batch)
                self.logger.info(f"Read values for batch {current_batch}: {values}")

                # Populate the dictionary
                for (tag, val, status, ts_string) in values:
                    formatted_ts = parse_timestamp(ts_string)
                    tag_values_map[tag] = (val, status, formatted_ts)

                # Write values to the original file
                write_values(self.filepath, tag_values_map)
                self.logger.info(f"Successfully wrote values for batch {current_batch} to {self.filepath}")

            except KeyboardInterrupt:
                self.logger.info("Stopping OPC Logger due to KeyboardInterrupt.")
                self.close_connection()
                sys.exit(0)
            except Exception as e:
                self.logger.error(f"Error during OPC read/write: {e}")
            finally:
                # Close OPC connection
                self.close_connection()

            if current_batch < total_batches:
                self.logger.info(f"Waiting for {self.interval} seconds before next batch.")
                try:
                    time.sleep(self.interval)
                except KeyboardInterrupt:
                    self.logger.info("Stopping OPC Logger during sleep due to KeyboardInterrupt.")
                    sys.exit(0)

            current_batch += 1

        self.logger.info("All batches processed. Exiting.")

@click.command()
@click.option(
    '--tagfile',
    type=click.Path(exists=True),
    required=True,
    help='Provide a CSV or XLSX file with a column called "Tag" and add the tags that need to be read from DeltaV / OPC Server.'
)
@click.option(
    '--servername',
    type=str,
    required=False,
    default="OPC.DeltaV.1",
    help='In case you wish to read the values from a different OPC server, you can provide the OPC server string here.'
)
@click.option(
    '--maxtagsperinterval',
    type=int,
    default=100,
    help='Provide a number that will be used to read the tags in burst mode.'
)
@click.option(
    '--intervalseconds',
    type=int,
    default=60,
    help='Provide a number in seconds that will be used to wait until the next burst runs.'
)
@click.option(
    '--disconnect_wait_time',
    type=int,
    default=10,
    help='Disconnect Wait time for the Server to Close the Connection.'
)
@click.option(
    '--info',
    is_flag=True,
    help='Display information with version about this tool.'
)
def main(tagfile, servername, maxtagsperinterval, intervalseconds, disconnect_wait_time, info):
    if info:
        click.echo(f"OPCLogger Tool Version {VERSION}\n{INFO_MESSAGE}")
        return

    if not tagfile:
        click.echo("Error: --tagfile is required input. Please provide a full path to the CSV or XLSX tag file.")
        sys.exit(1)

    logger = setup_logging()

    try:
        tags = read_tags(tagfile)
        logger.info(f"Successfully read {len(tags)} tags from {tagfile}")
    except Exception as e:
        logger.error(f"Failed to read tag file: {e}")
        sys.exit(1)

    opc_handler = OPCHandler(
        servername=servername,
        maxtags=maxtagsperinterval,
        interval=intervalseconds,
        disconnect_wait_time=disconnect_wait_time,
        tags=tags,
        filepath=tagfile,
        logger=logger
    )
    opc_handler.run()

if __name__ == '__main__':
    main()
