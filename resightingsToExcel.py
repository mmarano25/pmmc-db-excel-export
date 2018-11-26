"""Command line tool to convert resightings from a specific date range to Excel sheet"""

import argparse
import base64
import io
import json

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
    "Date": 2,
    "AnimalType": 3,
    "Latitude": 4,
    "Longitude": 5,
    "TagLocation": 6,
    "TagColor": 7,
    "DeadOrAlive": 8,
    "Condition": 9,
    "Injured": 10,
    "InjuredLocation": 11,
    "Entangled": 12,
    "EntangledLocation": 13,
    "NuisanceBehaviors": 14,
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

    # Add column headings
    for field_name, col_index in FIELDS.items():
        sheet.write(0, col_index, field_name)
        # Set starting column width
        sheet.set_column(0, col_index, 12)

    # For each resighting, slot the data in each field into its corresponding cell
    for row_num, resight in enumerate(resightings, 1):
        sheet.set_row(row_num, 175) # Set row height
        for field, data in resight.items():
            if field in FIELDS:
                # image data needs special handling
                # TODO: save image in a directory as well
                if "Image" in field: 
                    # Widen column here, if we do it earlier it's overriden
                    sheet.set_column(FIELDS[field], col_index, 45)

                    # Read from json as string
                    binary = io.BytesIO(base64.b64decode(data))
                    img = Image.open(binary)

                    # Resize to fit Excel
                    img.thumbnail((300, 300), Image.ANTIALIAS)  
                    resized = io.BytesIO()
                    img.save(resized, format=img.format)

                    # Insert into correct cell 
                    sheet.insert_image(
                        row_num, FIELDS[field], "image", {"image_data": resized}
                    )

                # Text data can be written as is
                else:
                    sheet.write(row_num, FIELDS[field], data)
            else:
                print("Field not recognized: {}".format(field))
    
    return workbook


if __name__ == "__main__":
    args = parser.parse_args()

    resightings = get_resightings_range(args.start, args.end)

    workbook = create_sheet_with_resightings(args.dest, resightings)

    workbook.close() # Saves workbook
