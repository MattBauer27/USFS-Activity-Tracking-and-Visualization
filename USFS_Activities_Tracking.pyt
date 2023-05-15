import os
import arcpy
import urllib.request
import zipfile
import openpyxl
import csv
from openpyxl.styles import Font, NamedStyle
import pandas as pd
import psycopg2
from psycopg2 import sql


class Toolbox(object):
    def __init__(self):
        self.label = "mkbauer_USCSSI_USFS_Activities_Tracking"
        self.alias = "mkbauer_USCSSI_USFS_Activities_Tracking"
        self.tools = [DownloadGeodatabases,
                      GenerateAcreageReport, ExportToPostgres]


class DownloadGeodatabases(object):
    def __init__(self):
        self.label = "Download USFS Activities Data"
        self.description = "Downloads geodatabases from URLs"
        self.canRunInBackground = False

    def getParameterInfo(self):
        # Input parameter for the output geodatabase
        output_gdb = arcpy.Parameter(
            displayName="Output Geodatabase",
            name="output_gdb",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input")

        params = [output_gdb]

        return params

    def execute(self, parameters, messages):

        # Get the output geodatabase path from the input parameter
        output_gdb = parameters[0].valueAsText

        # Create a reference to the current project
        aprx = arcpy.mp.ArcGISProject("CURRENT")

        # Set the output directory to the project's home folder
        output_dir = aprx.homeFolder

        # Set the URLs of the files to download and extract
        file_urls = [
            "https://data.fs.usda.gov/geodata/edw/edw_resources/fc/S_USA.Activity_RngVegImprove.gdb.zip",
            "https://data.fs.usda.gov/geodata/edw/edw_resources/fc/S_USA.Activity_SilvTSI.gdb.zip",
            "https://data.fs.usda.gov/geodata/edw/edw_resources/fc/S_USA.Activity_SilvReforestation.gdb.zip",
            "https://data.fs.usda.gov/geodata/edw/edw_resources/fc/S_USA.AdministrativeForest.gdb.zip",
            "https://data.fs.usda.gov/geodata/edw/edw_resources/fc/S_USA.AdministrativeRegion.gdb.zip"
        ]

        # Create an empty list to store the paths to the extracted geodatabases
        gdb_paths = []

        for file_url in file_urls:
            # Set the file name to download and extract
            file_name = os.path.basename(file_url)

            # Set the full path to the file to download and extract
            file_path = os.path.join(output_dir, file_name)

            # Download the file from the URL
            urllib.request.urlretrieve(file_url, file_path)

            # Extract the contents of the zip file to the output directory
            with zipfile.ZipFile(file_path, "r") as zip_ref:
                zip_ref.extractall(output_dir)

            # Delete the zip file
            os.remove(file_path)

            # Append the path to the extracted geodatabase to the list
            gdb_path = os.path.join(output_dir, file_name.replace(".zip", ""))
            gdb_paths.append(gdb_path)

        # Loop through the extracted geodatabases and copy their feature classes into the output geodatabase
        for gdb_path in gdb_paths:
            arcpy.env.workspace = gdb_path

            for fc_name in arcpy.ListFeatureClasses():
                output_fc_path = os.path.join(output_gdb, fc_name)
                arcpy.management.CopyFeatures(fc_name, output_fc_path)

            # Delete the extracted geodatabase
            arcpy.management.Delete(gdb_path)


class GenerateAcreageReport(object):

    def __init__(self):
        self.label = "Generate USFS Activities Acreage Report"
        self.description = "Creates an xlxs file that shows activies acreage, by year, for region, and forest"
        self.canRunInBackground = False

    def getParameterInfo(self):
        # Input feature layers
        act = arcpy.Parameter(
            name="act",
            displayName="Activity Data",
            datatype="GPFeatureLayer",
            parameterType="Required"
        )
        reg_bound = arcpy.Parameter(
            name="reg_bound",
            displayName="Regional Boundaries",
            datatype="GPFeatureLayer",
            parameterType="Required"
        )
        for_bound = arcpy.Parameter(
            name="for_bound",
            displayName="Forest Boundaries",
            datatype="GPFeatureLayer",
            parameterType="Required"
        )
        # Output Excel file location
        output_location = arcpy.Parameter(
            name="output_location",
            displayName="Output Location",
            datatype="DEFolder",
            parameterType="Required"
        )
        params = [act, reg_bound, for_bound, output_location]

        return params

    def execute(self, parameters, messages):
        # Get input parameters
        act = parameters[0].valueAsText
        reg_bound = parameters[1].valueAsText
        for_bound = parameters[2].valueAsText
        output_location = parameters[3].valueAsText

        if os.path.basename(act) == "Activity_RngVegImprove":
            act_name = "RangeVegetationImprovement"
        elif os.path.basename(act) == "Activity_SilvReforestation":
            act_name = "Reforestation"
        elif os.path.basename(act) == "Activity_SilvTSI":
            act_name = "TimberStandImprovement"

        workbook = openpyxl.Workbook()

        def BuildSheet(bound):

            if os.path.basename(bound) == 'AdministrativeRegion':
                boundary_name = 'REGIONNAME'
                sheet_name = 'Region'
            elif os.path.basename(bound) == 'AdministrativeForest':
                boundary_name = 'FORESTNAME'
                sheet_name = 'Forest'

            # Get a list of years from the FY_COMPLETED field in the three layers
            years = []
            with arcpy.da.SearchCursor(act, "FY_COMPLETED") as cursor:
                for row in cursor:
                    year = row[0]
                    if year not in years and year != None:
                        years.append(year)

            aprx = arcpy.mp.ArcGISProject("CURRENT")
            arcpy.management.CreateFileGDB(aprx.homeFolder, "TEMP")

            arcpy.analysis.Intersect([act, bound], os.path.join(
                aprx.homeFolder, "TEMP.gdb", os.path.basename(bound) + "_intersect"))

            arcpy.management.Dissolve(os.path.join(aprx.homeFolder, "TEMP.gdb", os.path.basename(bound) + "_intersect"), os.path.join(
                aprx.homeFolder, "TEMP.gdb", os.path.basename(bound) + "_dissolve"), ["FY_COMPLETED", boundary_name])

            arcpy.management.CalculateGeometryAttributes(os.path.join(aprx.homeFolder, "TEMP.gdb", os.path.basename(
                bound) + "_dissolve"), "ACRES AREA_GEODESIC", '', "ACRES_US", None, "SAME_AS_INPUT")

            arcpy.management.PivotTable(os.path.join(aprx.homeFolder, "TEMP.gdb", os.path.basename(
                bound) + "_dissolve"), boundary_name, "FY_COMPLETED", "ACRES", os.path.join(aprx.homeFolder, "TEMP.gdb", os.path.basename(bound) + "_pivoted"))

            if os.path.basename(bound) == "AdministrativeForest":
                arcpy.management.CalculateField(
                    in_table=os.path.join(
                        aprx.homeFolder, "TEMP.gdb", os.path.basename(bound) + "_pivoted"),
                    field="FORESTNAME",
                    expression='!FORESTNAME!.replace("Forests","").replace("Forest","").replace("National","")',
                    expression_type="PYTHON3"
                )
            elif os.path.basename(bound) == "AdministrativeRegion":
                arcpy.management.CalculateField(
                    in_table=os.path.join(
                        aprx.homeFolder, "TEMP.gdb", os.path.basename(bound) + "_pivoted"),
                    field="REGIONNAME",
                    expression='!REGIONNAME!.replace("Region","")',
                    expression_type="PYTHON3"
                )

            arcpy.conversion.TableToTable(os.path.join(aprx.homeFolder, "TEMP.gdb", os.path.basename(
                bound) + "_pivoted"), output_location, os.path.basename(bound) + "pivoted.csv")

            # read in the csv file
            with open(os.path.join(output_location, os.path.basename(bound) + "pivoted.csv"), 'r') as f:
                reader = csv.reader(f)
                data = [row for row in reader]

            # create a new list with the first column removed
            new_data = [row[1:] for row in data]

            # create a new worksheet and add the new data
            ws = workbook.create_sheet(sheet_name)
            for row in new_data:
                ws.append(row)

            sheet = workbook[sheet_name]

            # Loop through all cells in the first row
            for cell in sheet[1]:
                # Get the value of the cell
                cell_value = cell.value

                # Extract the 'FY_COMPLETED' part of the string
                if 'FY_COMPLETED' in cell_value:
                    fy_complet = cell_value[12:]
                elif 'REGIONNAME' in cell_value:
                    fy_complet = 'Region'
                elif 'FORESTNAME' in cell_value:
                    fy_complet = 'Forest'

                # Update the cell value with the extracted string
                cell.value = fy_complet

            arcpy.management.Delete(aprx.homeFolder + "/TEMP.gdb")
            arcpy.management.Delete(os.path.join(
                output_location, os.path.basename(bound) + "pivoted.csv"))

        BuildSheet(reg_bound)
        BuildSheet(for_bound)
        del workbook['Sheet']

        # Loop through each sheet in the workbook
        for sheet in workbook.worksheets:
            # Loop through each row in the sheet
            for row in sheet.rows:
                # Loop through each cell in the row
                for cell in row:
                    # Check if the cell contains a number
                    if isinstance(cell.value, (float, str)) and cell.value.replace('.', '').isdigit():
                        # Convert the number to an integer
                        cell.value = int(float(cell.value))

        # Loop through each sheet in the workbook
        for sheet in workbook.worksheets:
            # Auto adjust the column width to fit the longest value in each column
            for column in sheet.columns:
                max_length = 0
                column_name = openpyxl.utils.get_column_letter(
                    column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except TypeError:
                        pass
                adjusted_width = (max_length) * 1.2
                sheet.column_dimensions[column_name].width = adjusted_width

        # Define a named style for the bold font
        bold_style = NamedStyle(name='bold')
        bold_style.font = Font(bold=True)

        # Loop through each sheet in the workbook
        for sheet in workbook.worksheets:
            # Apply the bold style to the cells in the first row
            for cell in sheet[1]:
                cell.style = bold_style

        # save the changes to the workbook
        workbook.save(os.path.join(output_location,
                      act_name + "_AcreageReport.xlsx"))


class ExportToPostgres(object):
    def __init__(self):
        self.label = "Export Sheets in Workbook to PostgreSQL"
        self.description = "Exports sheets in Excel Workbook to PostgreSQL DB Table"
        self.canRunInBackground = False

    def getParameterInfo(self):
        # Input feature layers
        wb = arcpy.Parameter(
            name="wb",
            displayName="Workbook (.xlxs file)",
            datatype="DEFile",
            parameterType="Required"
        )

        db = arcpy.Parameter(
            name="db",
            displayName="Database Name",
            datatype="GPString",
            parameterType="Required"
        )

        user = arcpy.Parameter(
            name="user",
            displayName="Database Username",
            datatype="GPString",
            parameterType="Required"
        )

        pw = arcpy.Parameter(
            name="pw",
            displayName="Database Password",
            datatype="GPString",
            parameterType="Required"
        )

        host = arcpy.Parameter(
            name="host",
            displayName="Database Host",
            datatype="GPString",
            parameterType="Required"
        )

        port = arcpy.Parameter(
            name="port",
            displayName="Database Port",
            datatype="GPString",
            parameterType="Required"
        )

        sc = arcpy.Parameter(
            name="sc",
            displayName="Database Schema",
            datatype="GPString",
            parameterType="Required"
        )

        params = [wb, db, user, pw, host, port, sc]

        return params

    def execute(self, parameters, messages):
        # Get input parameters
        wb = parameters[0].valueAsText
        db = parameters[1].valueAsText
        user = parameters[2].valueAsText
        pw = parameters[3].valueAsText
        host = parameters[4].valueAsText
        port = parameters[5].valueAsText
        sc = parameters[6].valueAsText

        # Read excel workbook
        wb_open = openpyxl.load_workbook(wb)
        sheetnames = wb_open.sheetnames

        # Get the workbook name without the file extension
        workbook_name = os.path.splitext(os.path.basename(wb))[0]

        # Connect to your PostgreSQL database
        conn = psycopg2.connect(database=db, user=user,
                                password=pw, host=host, port=port)
        cur = conn.cursor()

        for sheet in sheetnames:
            # Read sheet as a dataframe
            df = pd.read_excel(wb_open, sheet_name=sheet, engine='openpyxl')

            # Replace spaces with underscores in column names, prepend 'yr_' to numeric column names
            df.columns = [f"yr_{str(col)}" if isinstance(
                col, (int, float)) else col.replace(' ', '_') for col in df.columns]

            # Combine workbook name and sheet name for the table name
            table_name = f"{workbook_name}_{sheet}".lower()

            # Check if the table exists
            cur.execute("SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = %s AND table_name=%s)", (sc, table_name))
            if cur.fetchone()[0]:
                # If the table exists, drop the table
                cur.execute(f"DROP TABLE {sc}.{table_name}")

            # Create the table
            columns_definition = [f"{df.columns[0]} text"] + \
                [f"{col} double precision" for col in df.columns[1:]]
            columns = ", ".join(columns_definition)
            create_table_query = f"CREATE TABLE {sc}.{table_name} ({columns})"
            cur.execute(create_table_query)

            # Insert new data
            for index, row in df.iterrows():
                # Convert non-integer columns to float
                row_values = [
                    row[0]] + [float(val) if not pd.isna(val) else None for val in row[1:]]

                columns = ', '.join(df.columns)
                placeholders = ', '.join(['%s' for col in df.columns])
                insert_query = f"INSERT INTO {sc}.{table_name} ({columns}) VALUES ({placeholders})"
                cur.execute(insert_query, tuple(row_values))

            # Commit the changes
            conn.commit()

        # Close the cursor and the connection
        cur.close()
        conn.close()
