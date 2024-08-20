using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LifestyleDesign_r24.Classes;
using LifestyleDesign_r24.Common;

namespace LifestyleDesign_r24
{
    [Transaction(TransactionMode.Manual)]
    public class cmdShakeAreaBoundary : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document curDoc = uiapp.ActiveUIDocument.Document;

            // get all the area plans in the project
            List<View> areaViews = new List<View>();

            List<View> areaFloor = Utils.GetAllViewsByCategory(curDoc, "10:Floor Areas");
            List<View> areaFrame = Utils.GetAllViewsByCategory(curDoc, "11:Frame Areas");
            List<View> areaAttic = Utils.GetAllViewsByCategory(curDoc, "12:Attic Areas");

            areaViews.AddRange(areaFloor);
            areaViews.AddRange(areaFrame);
            areaViews.AddRange(areaAttic);

            // start the transaction
            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Shake Area Boundary Lines");
                {
                    // loop through each view in the list
                    foreach (View curView in areaViews)
                    {
                        // get all area boundary lines in the active view
                        IEnumerable<Element> colABLines = new FilteredElementCollector(curDoc, curView.Id)
                            .OfCategory(BuiltInCategory.OST_AreaSchemeLines)
                            .WhereElementIsNotElementType();

                        // get all the area tags
                        FilteredElementCollector colAreaTags = new FilteredElementCollector(curDoc, curView.Id)
                            .OfClass(typeof(SpatialElementTag));

                        foreach (SpatialElementTag curTag in colAreaTags)
                        {
                            curTag.HasLeader = false;
                        }

                        if (colABLines.Count() > 0)
                        {
                            // get the first line in the list
                            Element lineToMove = colABLines.FirstOrDefault();

                            // get the location of the line
                            LocationCurve curLocation = lineToMove.Location as LocationCurve;

                            // create a vector to move the line
                            XYZ curPoint = curLocation.Curve.GetEndPoint(0) as XYZ;
                            XYZ newVector = new XYZ(.25, 0, 0);
                            XYZ oldVector = new XYZ( -.25, 0,0);

                            // move the line to the left
                            lineToMove.Location.Move(newVector);

                            // move the line back to the right                
                            lineToMove.Location.Move(oldVector);
                        }
                        else
                            continue;
                    }
                }

                t.Commit();
            }

            return Result.Succeeded;            
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand5_1";
            string buttonTitle = "Shake\rBoundaries";
            string methodBase = MethodBase.GetCurrentMethod().DeclaringType?.FullName;

            if (methodBase == null)
            {
                throw new InvalidOperationException("MethodBase.GetCurrentMethod().DeclaringType?.FullName is null");
            }
            else
            {
                clsButtonData myButtonData1 = new Classes.clsButtonData(
                    buttonInternalName,
                    buttonTitle,
                    methodBase,
                    Properties.Resources.Green_32,
                    Properties.Resources.Green_16,
                    "Shifts area boundary lines back and forth to correct negative numbers");

                return myButtonData1.Data;
            }
        }
    }
}

