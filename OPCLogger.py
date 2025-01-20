import click
import pandas as pd
import OpenOPC
import time
import logging
import sys

VERSION = "1.0.0"

INFO_MESSAGE = """
This tool was designed to run an OPC Reader utility to get values from DeltaV OPC Server in burst modes.
It then saves the data into the same CSV/XLSX file it was provided with. The input file must contain a column named "Tag".

## Argument Requirements

- `--tagfile`: Provide a CSV or XLSX file with a column called "Tag" and add the tags that need to be read from DeltaV / OPC Server.

- `--servername`: In case you wish to read the values from a different OPC server, you can provide the OPC server string here.

- `--maxtagsperinterval`: Provide a number that will be used to read the tags in burst mode.

- `--intervalseconds`: Provide a number in seconds that will be used to wait until the next burst runs.
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

def write_values(filepath, values):
    """
    Writes the provided list of `values` into the same file as a 'Value' column.
    The file must already exist with at least the 'Tag' column. 
    """
    try:
        if filepath.endswith('.xlsx'):
            # Read existing Excel file
            df = pd.read_excel(filepath, engine='openpyxl')
            # Add or update the 'Value' column
            df['Value'] = values
            # Write back to Excel
            df.to_excel(filepath, index=False, engine='openpyxl')
        elif filepath.endswith('.csv'):
            # Read existing CSV file
            df = pd.read_csv(filepath)
            # Add or update the 'Value' column
            df['Value'] = values
            # Write back to CSV
            df.to_csv(filepath, index=False)
        else:
            raise ValueError("Unsupported file format. Please provide a .xlsx or .csv file.")
    except Exception as e:
        raise ValueError(f"Error writing to the tag file: {e}")

class OPCHandler:
    def __init__(self, servername, maxtags, interval, tags, filepath, logger):
        self.servername = servername
        self.maxtags = maxtags
        self.interval = interval
        self.tags = tags
        self.filepath = filepath
        self.logger = logger
        self.opc = OpenOPC.client()

    def connect(self):
        try:
            self.opc.connect(self.servername)
            self.logger.info(f"Connected to OPC server: {self.servername}")
        except Exception as e:
            self.logger.error(f"Failed to connect to OPC server: {e}")
            sys.exit(1)

    def run(self):
        self.connect()
        while True:
            try:
                all_values = []
                for i in range(0, len(self.tags), self.maxtags):
                    batch = self.tags[i:i+self.maxtags]
                    values = self.opc.read(batch)
                    self.logger.info(f"Read values for batch {i//self.maxtags + 1}: {values}")
                    all_values.extend(values)
                write_values(self.filepath, all_values)
                self.logger.info(f"Successfully wrote values to {self.filepath}")
                time.sleep(self.interval)
            except KeyboardInterrupt:
                self.logger.info("Stopping OPC Logger.")
                self.opc.close()
                sys.exit(0)
            except Exception as e:
                self.logger.error(f"Error during OPC read/write: {e}")
                time.sleep(self.interval)

@click.command()
@click.option(
    '--tagfile',
    type=click.Path(exists=True),
    required=False,
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
    '--info',
    is_flag=True,
    help='Display information with version about this tool.'
)
def main(tagfile, servername, maxtagsperinterval, intervalseconds, info):
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

    opc = OPCHandler(
        servername=servername,
        maxtags=maxtagsperinterval,
        interval=intervalseconds,
        tags=tags,
        filepath=tagfile,
        logger=logger
    )
    opc.run()

if __name__ == '__main__':
    main()
