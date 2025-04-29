using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

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

                    #region First Part Of Task

                    
                    // Create a new structural column
                    FamilySymbol columnType = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_StructuralColumns)
                        .FirstOrDefault( Type => Type.Name == "600 x 750mm") as FamilySymbol;

                    // Check if the column type is valid
                    if (columnType != null && !columnType.IsActive)
                    {
                        columnType.Activate();
                        doc.Regenerate();
                    }

                    // Create a new structural column at the origin point
                    

                    List<XYZ> points = new List<XYZ>();

                    points.Add(new XYZ(0, 0, 0));   
                    points.Add(new XYZ(5, 0, 0));  
                    points.Add(new XYZ(10, 0, 0));   
                    points.Add(new XYZ(15, 0, 0));   
                    points.Add(new XYZ(20, 0, 0));


                    // Find a level to place the column on

                    Level selectedLevel = new FilteredElementCollector(doc)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .FirstOrDefault(l => l.Name == "Level 1")
                        ?? new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .FirstOrDefault(l => l.Name == "Level 2")
                        ?? new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .FirstElement() as Level;

                    if (selectedLevel == null)
                    {
                        using (var trans = new SubTransaction(doc))
                        {
                            trans.Start();


                            selectedLevel = Level.Create(doc, 0);
                            selectedLevel.Name = "Level 1";

                            trans.Commit();
                        }

                        TaskDialog.Show( "info" , "Level 1 has been created succesfully");
                    }

                    List<FamilyInstance> newlyCreatedColumns = new List<FamilyInstance>();

                    foreach (XYZ point in points)
                    {
                        // Create the column at the specified point
                        FamilyInstance column = doc.Create.NewFamilyInstance(point, columnType, selectedLevel, StructuralType.Column);

                        newlyCreatedColumns.Add(column);

                    }
                    #endregion

                    #region Second Part Of Task 

                    List<string> OldComments = new List<string> { "A", "B", "C", "D" , "E"};

                    using (var commenttrans= new SubTransaction(doc))
                    {
                        commenttrans.Start();
                        foreach (FamilyInstance column in newlyCreatedColumns)
                        {
                            // Get the parameter for comments
                            Parameter commentParam = column.LookupParameter("Comments");
                            // Check if the parameter is valid
                            if (commentParam != null && !commentParam.IsReadOnly)
                            {
                                // Set the comment value
                                int index = newlyCreatedColumns.IndexOf(column);
                                commentParam.Set(OldComments[index]);
                            }
                        }
                        commenttrans.Commit();

                    }

                    // Show final results
                    TaskDialog.Show("Success",
                        $"Processed {Math.Min(newlyCreatedColumns.Count, OldComments.Count)} columns\n\n" +
                        "Click OK to view details",
                        TaskDialogCommonButtons.Ok);

                    // Show details when OK is clicked
                    string details = "Column Details:\n";
                    for (int i = 0; i < newlyCreatedColumns.Count && i < OldComments.Count; i++)
                    {
                        FamilyInstance column = newlyCreatedColumns[i];
                        Parameter commentParam = column.LookupParameter("Comments");
                        string actualComment = commentParam?.AsString() ?? "No comment assigned";

                        details += $"Column {i + 1}: => Comment {actualComment}\n";
                    }

                    TaskDialog.Show("Comment Details", details);

                    List<string> NewComments = new List<string> { "V", "W", "X", "Y", "Z" };

                    using (var updateCommentTrans= new SubTransaction(doc))
                    {
                        updateCommentTrans.Start();

                        foreach (FamilyInstance column in newlyCreatedColumns)
                        {
                            // Get the parameter for comments
                            Parameter commentParam = column.LookupParameter("Comments");
                            // Check if the parameter is valid
                            if (commentParam != null && !commentParam.IsReadOnly)
                            {
                                // Set the comment value
                                int index = newlyCreatedColumns.IndexOf(column);
                                commentParam.Set(NewComments[index]);
                            }
                        }

                        updateCommentTrans.Commit();
                    }

                    // Show final results
                    TaskDialog.Show("Success",
                        $"Processed {Math.Min(newlyCreatedColumns.Count, OldComments.Count)} columns\n\n" +
                        "Click OK to view updated details",
                        TaskDialogCommonButtons.Ok);

                    // Show details when OK is clicked
                    string updatedDetails = "Column Details:\n";

                    for (int i = 0; i < newlyCreatedColumns.Count && i < OldComments.Count; i++)
                    {
                        FamilyInstance column = newlyCreatedColumns[i];
                        Parameter commentParam = column.LookupParameter("Comments");
                        string actualComment = commentParam?.AsString() ?? "No comment assigned";

                        updatedDetails += $"Column {i + 1}: => Comment {actualComment}\n";
                    }

                    TaskDialog.Show("Comment Details", updatedDetails);





                    #endregion
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
