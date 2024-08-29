using LifestyleDesign_r24.Classes;
using LifestyleDesign_r24.Common;

namespace LifestyleDesign_r24
{
    [Transaction(TransactionMode.Manual)]
    public class cmdUpdateVTs : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document curDoc = uiapp.ActiveUIDocument.Document;

            // this is a variable for the current Revit model in the UI
            UIDocument uidoc = uiapp.ActiveUIDocument;

            // get all the view templates in the project
            List<View> curVTs = Utils.GetAllViewTemplates(curDoc);

            // set the path to the view template file
            string templateDoc = "S:\\Shared Folders\\Lifestyle USA Design\\Library 2025\\Template\\View Templates.rvt";

            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Update View Templates");

                // delete all view templates that start with a letter or a number, except 17
                foreach (View curVT in curVTs)
                {
                    // get the name of the view template
                    string curName = curVT.Name;

                    // check if first character is letter
                    bool isLetter = !String.IsNullOrEmpty(curName) && Char.IsLetter(curName[0]);

                    // check if first two charactera is number
                    // string firstTwo = curName.Substring(0, 1);
                    bool isNumber = !String.IsNullOrEmpty(curName) && Char.IsNumber(curName[0]);


                    // if yes, delete it
                    if (isLetter == true || isNumber == true)
                    {
                        try
                        {
                            curDoc.Delete(curVT.Id);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }

                // transfer view templates from template file
                Document sourceDoc = uidoc.Application.OpenAndActivateDocument(templateDoc).Document;
                Document targetDoc = uidoc.Document;

                // get the view templates from the source document
                List<View> listViewTemplates = new FilteredElementCollector(sourceDoc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => v.IsTemplate)
                    .ToList();

                foreach (View sourceTemplate in listViewTemplates)
                {
                    // duplicate the view template in the target document
                    ElementId newTemplateID = Utils.DuplcateViewTemplate(sourceDoc, sourceTemplate, targetDoc);
                }

                t.Commit();
            }

            // transfer the current view templates from the template file

            // assign view templates to views


            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2_1";
            string buttonTitle = "Update\rView Templates";
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
                    Properties.Resources.MirrorProject_32,
                    Properties.Resources.MirrorProject_16,
                    "Updates View Templates to current standards and applies them to the views");

                return myButtonData1.Data;
            }
        }
    }
}
