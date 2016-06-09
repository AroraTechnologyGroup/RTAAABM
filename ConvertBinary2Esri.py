import os
import base64
import arcpy
from openpyxl import Workbook, load_workbook
from arcpy import env, Geometry, da
import traceback
import sys

target_folder = env.scratchFolder
env.workspace = target_folder
sr = arcpy.SpatialReference(4326)
env.overwriteOutput = 1


def process_table(table, gdb, dset):
    out_dataset = "{}\\{}".format(gdb, dset)
    env.workspace = out_dataset
    in_table = table
    wb = load_workbook(filename=in_table)
    ws_name = wb.get_sheet_names()[0]
    ws = wb['{}'.format(ws_name)]

    polygon = arcpy.ListFeatureClasses("Polygon*")[0]
    polyline = arcpy.ListFeatureClasses("Polyline*")[0]
    point = arcpy.ListFeatureClasses("Point*")[0]

    type_dict = {"polygon": polygon, "polyline": polyline, "point": point}
    for sh_type in type_dict.keys():
        cursor = da.InsertCursor(type_dict[sh_type], ["SHAPE@"])
        if len(ws.rows):
            for row in ws.rows:
                entry = ""
                if len(row):
                    entry = row[0]
                if entry != "":
                    array = bytearray(base64.b64decode(entry.value))
                    # we don't know the geometry type until after it is imported
                    in_geo = ""
                    try:
                        in_geo = arcpy.FromWKB(array)

                        geo_type = in_geo.type
                        if geo_type == sh_type:
                            try:
                                cursor.insertRow([in_geo])
                            except:
                                exc_type, exc_value, exc_traceback = sys.exc_info()
                                print "{} :: {}".format(traceback.print_tb(exc_traceback, limit=1, file=sys.stdout),
                                                        traceback.print_exception(exc_type, exc_value, exc_traceback,
                                                                                  limit=2, file=sys.stdout))
                    except:
                        exc_type, exc_value, exc_traceback = sys.exc_info()
                        print "{} :: {}".format(traceback.print_tb(exc_traceback, limit=1, file=sys.stdout),
                                                traceback.print_exception(exc_type, exc_value, exc_traceback,
                                                                          limit=2, file=sys.stdout))
        del cursor


def process_folder(path, gdb):
    i = 0
    excel_files = os.listdir(path)
    for fl in excel_files:
        try:
            env.workspace = gdb
            form_name = fl.split(".")[0].replace(" ", "_")
            arcpy.CreateFeatureDataset_management(gdb, form_name, sr)
            env.workspace = "{}\\{}".format(gdb, form_name)
            for shape in ["Polygon", "Polyline", "Point"]:
                fc_name = "{}{}".format(shape, i)
                if arcpy.Exists(fc_name):
                    arcpy.Delete_management(fc_name)
                arcpy.CreateFeatureclass_management("{}\\{}".format(gdb, form_name),
                                                    fc_name, shape.upper())
                i += 1
            xl_path = "{}\\{}".format(path, fl)
            process_table(xl_path, gdb, form_name)

        except:
            exc_type, exc_value, exc_traceback = sys.exc_info()
            print "{} :: {}".format(traceback.print_tb(exc_traceback, limit=1, file=sys.stdout),
                                    traceback.print_exception(exc_type, exc_value, exc_traceback,
                                                              limit=2, file=sys.stdout)
                                    )
folder_path = r"C:\RTAA_ABM"
os.chdir(folder_path)
folders = os.listdir(folder_path)
folders = [f for f in folders if os.path.isdir(f)]
print folders

for x in folders:
    in_path = "{}\\{}".format(folder_path, x)
    gdb_name = x.replace(" ", "_")
    gdb_path = "{}\\{}.gdb".format(env.scratchFolder, gdb_name)
    if not arcpy.Exists(gdb_path):
        arcpy.CreateFileGDB_management(env.scratchFolder, gdb_name)

    process_folder(in_path, gdb_path)




