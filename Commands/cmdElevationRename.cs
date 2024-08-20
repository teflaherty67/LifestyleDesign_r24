using LifestyleDesign_r24.Classes;
using LifestyleDesign_r24.Common;

namespace LifestyleDesign_r24
{
    [Transaction(TransactionMode.Manual)]
    public class cmdElevationRename : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document curDoc = uiapp.ActiveUIDocument.Document;

            // get all the elevation views
            List<View> viewList = Utils.GetAllElevationViews(curDoc);

            // create an empty list for renamed views
            List<View> renamedList = new List<View>();

            // create a counter for informing the user
            int counter = 0;

            // start the transaction
            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Rename Elevations");

                // loop through the view list
                foreach (View curView in viewList)
                {
                    Parameter titleParam = null;

                    string curTitle = "";

                    foreach (Parameter curParam in curView.Parameters)
                    {
                        if (curParam.Definition.Name == "Title on Sheet")
                        {
                            titleParam = curParam;

                            curTitle = curParam.AsString();
                        }
                    }

                    // change view name
                    if (curView.Name.Contains("Left") == true)
                    {
                        curView.Name = curView.Name.Replace("Left", "$Right");
                        renamedList.Add(curView);
                    }

                    else if (curView.Name.Contains("Right") == true)
                    {
                        curView.Name = curView.Name.Replace("Right", "$Left");
                        renamedList.Add(curView);
                    }

                    // change the title on sheet

                    if (curTitle.Contains("Left"))
                        titleParam.Set(curTitle.Replace("Left", "Right"));

                    else if (curTitle.Contains("Right"))
                        titleParam.Set(curTitle.Replace("Right", "Left"));
                }

                foreach (View curView in renamedList)
                {
                    // remove $ from view name
                    if (curView.Name.Contains("$Right") == true)
                        curView.Name = curView.Name.Replace("$Right", "Right");

                    else if (curView.Name.Contains("$Left") == true)
                        curView.Name = curView.Name.Replace("$Left", "Left");

                    // increment the counter
                    counter++;
                }

                // alert the user
                string msgText = "Renamed " + counter.ToString() + " views.";
                string msgTitle = "Complete";
                Forms.MessageBoxButton msgButtons = Forms.MessageBoxButton.OK;

                Forms.MessageBox.Show(msgText, msgTitle, msgButtons, Forms.MessageBoxImage.Information);

                // commit the changes

                t.Commit();

                return Result.Succeeded;
            }
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand3_1";
            string buttonTitle = "Rename\rElevations";
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
                    Properties.Resources.ElevationRename_32,
                    Properties.Resources.ElevationRename_16,
                    "Renames Left & Right Elevations for Mirrored Projects");

                return myButtonData1.Data;
            }
        }
    }
}
