"""
Script to get resightings from the database between a specified date range and export 
them to a formatted Excel sheet with copies of each resighting in their own folder.

- It can be used either as a command line tool like or the resightings_to_Excel function
  can be imported and used directly.
- Note that full sized copies of the records are saved separately because the Excel 
  library used does not support scaling images to fit the cells, only resizing them. 
"""

import argparse
import base64
import io
import json
import datetime

import boto3
from boto3.dynamodb.conditions import Attr
from PIL import Image
from xlsxwriter import Workbook


# Command line setup. The description is the first few lines of the docstring.
parser = argparse.ArgumentParser(description="\n".join(__doc__.splitlines()[0:3])) 
parser.add_argument("start", help="Range start date.")
parser.add_argument("end", help="Range end date.")
parser.add_argument("dest", help="Path to save Excel sheet at.")

# Database information
REGION_NAME = 'us-west-2'
ENDPOINT_URL = "http://localhost:8000"
TABLE_NAME = "Resightings"

# Fields and their corresponding column in the final Excel
FIELDS = {
    "SituationImage": 0,
    "TagImage": 1,
    "Timestamp": 2,
    "AnimalType": 3,
    "Latitude": 4,
    "Longitude": 5,
    "Location": 6,
    "TagLocation": 7,
    "TagColor": 8,
    "DeadOrAlive": 9,
    "Condition": 10,
    "Injured": 11,
    "Entangled": 12,
    "InjuredOrEntangledLocation": 13,
    "NuisanceBehaviors": 14,
}


def get_resightings_range(date1: str, date2: str) -> [dict]:
    """Pulls desired resightings from cloud database"""
    try:
        # Confirm date is in correct format
        date1_epoch = int(datetime.datetime.strptime(date1, "%m/%d/%Y").strftime("%s"))
        date2_epoch = int(datetime.datetime.strptime(date2, "%m/%d/%Y").strftime("%s"))
    except ValueError:
        raise RuntimeError("Could not read date format. Enter as day/month/year")

    # Query database over date range
    dynamodb = boto3.resource('dynamodb', region_name=REGION_NAME, endpoint_url=ENDPOINT_URL)
    table = dynamodb.Table(TABLE_NAME)

    result = table.scan(
        FilterExpression=Attr("Timestamp").between(
            date1_epoch, date2_epoch)
    )

    return result["Items"]


def create_directory_of_resightings(resightings: list, dir_name: str) -> None:
    # |- Resightings from Date1 to Date2
    # |- |- Excel sheet
    # |- |- ResightingDateXAnimalTypeYGPSLatLong
    # |- |- |- SituationImage.jpg
    # |- |- |- TagImage.jpg
    # |- |- -- FormData.txt
    # |- |- ResightingDateXAnimalTypeYGPSLatLong
    # |- |- |- SituationImage.jpg
    # |- |- |- TagImage.jpg
    # |- -- -- FormData.txt
    # TODO: 
    #   - Pretty print form data in text file
    #   - Create dir with correct name
    pass


def create_sheet_with_resightings(workbook_name: str, resightings: list) -> Workbook:
    """Places resightings in workbook with named columns, treating images specially"""
    if not resightings:
        return
    
    workbook = Workbook(workbook_name)
    sheet = workbook.add_worksheet()

    # Format object to center data and wrap text
    cell_format = workbook.add_format()
    cell_format.set_shrink()
    cell_format.set_text_wrap()
    cell_format.set_align("center")
    cell_format.set_align("vcenter")

    # Add column headings
    for field_name, col_index in FIELDS.items():
        sheet.write(0, col_index, field_name)
        # Set starting column width (modifies from range col_index -> col_index)
        if "Image" in field_name:
            # Images need a wider column
            sheet.set_column(col_index, col_index, 45, cell_format)
        else:
            sheet.set_column(col_index, col_index, 12, cell_format)


    # For each resighting, slot the data in each field into its corresponding cell
    for row_num, resight in enumerate(resightings, 1):
        sheet.set_row(row_num, 160) # Set row height
        for field, data in resight.items():
            if field in FIELDS:
                # image data needs special handling
                # TODO: save image in a directory as well
                if "Image" in field: 
                    # Read from json as string
                    binary = io.BytesIO(base64.b64decode(data))
                    img = Image.open(binary)

                    # Resize to fit Excel
                    img.thumbnail((300, 300), Image.ANTIALIAS)  
                    resized = io.BytesIO()
                    img.save(resized, format=img.format)

                    # Insert into correct cell 
                    sheet.insert_image(
                        row_num, FIELDS[field], "image", {"image_data": resized, "positioning": 1}
                    )

                # Write other data, converting if necessary
                else:
                    if field == "Timestamp":
                        # Convert epoch time to month/day/year
                        data = datetime.datetime.fromtimestamp(data).strftime("%m/%d/%Y %H:%M:%S")

                    sheet.write(row_num, FIELDS[field], data)
            else:
                print("Field not recognized: {}".format(field))
    
    return workbook

def get_resightings_to_Excel(start: str, end: str, dest: str) -> None:
    resightings = get_resightings_range(start, end)

    workbook = create_sheet_with_resightings(dest, resightings)

    workbook.close() # Saves workbook


if __name__ == "__main__":
    args = parser.parse_args()
    get_resightings_to_Excel(args.start, args.end, args.dest)
