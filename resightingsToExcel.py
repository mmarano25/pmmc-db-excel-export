"""Command line tool to convert resightings from a specific date range to Excel sheet"""

import argparse
import base64
import io
import json
import datetime

import boto3
from PIL import Image
from xlsxwriter import Workbook


# Command line setup
parser = argparse.ArgumentParser(
    description="Gets sightings from cloud database from date1 to date2 and exports "
                "them into an excel sheet saved at dest."
)
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
    "Location": 4,
    "TagLocation": 5,
    "TagColor": 6,
    "DeadOrAlive": 7,
    "Condition": 8,
    "Injured": 9,
    "InjuredLocation": 10,
    "Entangled": 11,
    "EntangledLocation": 12,
    "NuisanceBehaviors": 13,
}


def get_resightings_range(date1: str, date2: str) -> [dict]:
    """Pulls desired resightings from cloud database"""
    # Currently just gets all in table
    dynamodb = boto3.resource('dynamodb', region_name=REGION_NAME, endpoint_url=ENDPOINT_URL)
    table = dynamodb.Table(TABLE_NAME)
    data = table.scan()
    return data["Items"]


def create_sheet_with_resightings(workbook_name: str, resightings: list) -> Workbook:
    """Places resightings in workbook with named columns, treating images specially"""
    if not resightings:
        return
    
    workbook = Workbook(workbook_name)
    sheet = workbook.add_worksheet()

    # Format object to center data and wrap text
    cell_format = workbook.add_format()
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
        sheet.set_row(row_num, 175) # Set row height
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


if __name__ == "__main__":
    args = parser.parse_args()

    resightings = get_resightings_range(args.start, args.end)

    workbook = create_sheet_with_resightings(args.dest, resightings)

    workbook.close() # Saves workbook
