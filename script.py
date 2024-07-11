#region library
import clr 
import os
import sys
clr.AddReference("System")
import System

clr.AddReference("RevitServices")
import RevitServices
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Collections.Generic import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *


from System.Windows import MessageBox
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Markup import XamlReader
#endregion

#region revit infor
# Get the directory path of the script.py & the Window.xaml
dir_path = os.path.dirname(os.path.realpath(__file__))
#xaml_file_path = os.path.join(dir_path, "Window.xaml")

#Get UIDocument, Document, UIApplication, Application
uidoc = __revit__.ActiveUIDocument
uiapp = UIApplication(uidoc.Document.Application)
app = uiapp.Application
doc = uidoc.Document
activeView = doc.ActiveView

#general infor 
all_FillPatterns = FilteredElementCollector(doc).OfClass(FillPatternElement).WhereElementIsNotElementType().ToElements()
all_grids = FilteredElementCollector(doc).OfClass(Grid).WhereElementIsNotElementType().ToElements()
magenta_color = Color(255, 0, 255)
yellow_color = Color(255, 250, 0)
PATTERN_NAME = "<Solid fill>"

#endregion

#region method
class Utils:

    def highlight_color(self, columnList, color):
        setting = OverrideGraphicSettings()
        setting.SetCutForegroundPatternColor(color)
        setting.SetCutBackgroundPatternColor(color)
        setting.SetSurfaceBackgroundPatternColor(color)
        setting.SetSurfaceForegroundPatternColor(color)

        for pattern in all_FillPatterns:
            if pattern.Name == PATTERN_NAME:
                setting.SetCutBackgroundPatternId(pattern.Id)
                setting.SetCutForegroundPatternId(pattern.Id)
                setting.SetSurfaceBackgroundPatternId(pattern.Id)
                setting.SetSurfaceForegroundPatternId(pattern.Id)
                break

        for column in columnList:
            activeView.SetElementOverrides(column.Id, setting)

    def reset_color(self, columnList):
        setting = OverrideGraphicSettings()
        for column in columnList:
            activeView.SetElementOverrides(column.Id, setting)
            

    def check_colums_location(self, columnsList, picked_column):
        x = picked_column.Location.Point.X
        y = picked_column.Location.Point.Y

        columns_to_highlight = []
        columns_to_reset = []

        for column in columnsList:
            x1 = (column.Location.Point.X - x) * 304.8
            y1 = (column.Location.Point.Y - y) * 304.8
            
            if round(x1, 0) != 0 or round(y1, 0) != 0:
                columns_to_highlight.append(column)
            else: 
                columns_to_reset.append(column)

        self.highlight_color(columns_to_highlight, magenta_color)
        self.reset_color(columns_to_reset)

        return columns_to_highlight


#endregion
class FilterColumn(ISelectionFilter):
    def AllowElement(self, element):
        if element.Category.Name == "Structural Columns": return True
        else: return False
             
    def AllowReference(self, reference, position):
        return True
    
#select elements
class Main ():
    def main_task(self):
        
        #select and pick column in active view
        selected_objects = []
        picked_column = None
        try:
            selected_objects = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, FilterColumn())
            if selected_objects is not None and len(selected_objects) > 0:
                picked_ele = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, FilterColumn())
                picked_column = doc.GetElement(picked_ele)
        except:
            pass

        if len(selected_objects) > 0 and picked_column is not None:

            #classify vertical and slanted
            vertical_columns = []
            slanted_columns = []

            for r in selected_objects:
                column = doc.GetElement(r)
                index = column.get_Parameter(BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM).AsInteger()
                if index == 0:
                    vertical_columns.append(column)
                else:
                    slanted_columns.append(column)
        
            #check column location
            total = 0
            try:

                t = Transaction(doc, " ")
                t.Start()

                checked_columns = Utils().check_colums_location(vertical_columns, picked_column)
                Utils().highlight_color(slanted_columns, yellow_color)
                total = len(checked_columns) + len (slanted_columns)

                t.Commit()
            except Exception as e:
                MessageBox.Show(str(e), "Message")

            message = "Found "+ str(total) + " columns that need to check the position!"
            MessageBox.Show(message, "Message")
        

if __name__ == "__main__":
    Main().main_task()
                
    
    





