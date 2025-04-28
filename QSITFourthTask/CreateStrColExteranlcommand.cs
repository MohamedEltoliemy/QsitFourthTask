using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;

namespace QSITFourthTask
{
    [Transaction(TransactionMode.Manual)]
    public class CreateStrColExteranlcommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get the active document
                UIApplication uiapp = commandData.Application;
                Document doc = commandData.Application.ActiveUIDocument.Document;
               using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Structural Column");
                    // Create a new structural column
                    FamilySymbol columnType = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_StructuralColumns)
                        .FirstElement() as FamilySymbol;

                    // Check if the column type is valid
                    if (columnType != null && !columnType.IsActive)
                    {
                        columnType.Activate();
                        doc.Regenerate();
                    }

                    // Create a new structural column at the origin point
                    XYZ point = new XYZ(0, 0, 0);
                    Level level = new FilteredElementCollector(doc)
                        .OfClass(typeof(Level))
                        .FirstElement() as Level;
                    FamilyInstance column = doc.Create.NewFamilyInstance(point, columnType,level, StructuralType.Column);
                    t.Commit();
                }



                return Result.Succeeded;
            }
            catch(Exception e )
            {
                TaskDialog.Show("Error", e.Message);
                return Result.Failed;
            }
           
        }
    }
}
