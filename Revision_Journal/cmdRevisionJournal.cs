using LifestyleDesign_r24.Classes;

namespace LifestyleDesign_r24
{
    [Transaction(TransactionMode.Manual)]
    public class cmdRevisionJournal : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document curDoc = uiapp.ActiveUIDocument.Document;

            frmJobNumber curForm1 = new frmJobNumber();

            curForm1.Width = 375;
            curForm1.Height = 150;

            curForm1.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
            curForm1.Topmost = true;

            curForm1.ShowDialog();

            frmRevisionJournal curForm = new frmRevisionJournal(curDoc.PathName);

            curForm.Width = 700;
            curForm.Height = 500;

            curForm.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
            curForm.Topmost = true;

            curForm.ShowDialog();

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";
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
                    Properties.Resources.Blue_32,
                    Properties.Resources.Blue_16,
                    "This is a tooltip for Button 1");

                return myButtonData1.Data;
            }
        }
    }
}
