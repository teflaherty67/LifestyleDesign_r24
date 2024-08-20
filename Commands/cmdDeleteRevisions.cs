using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LifestyleDesign_r24.Classes;

namespace LifestyleDesign_r24
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDeleteRevisions : IExternalCommand
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

            string msgText = "About to delete all Revisions.";
            string msgTitle = "Warning";
            Forms.MessageBoxButton msgButtons = Forms.MessageBoxButton.OKCancel;

            Forms.MessageBoxResult result = Forms.MessageBox.Show(msgText, msgTitle, msgButtons, Forms.MessageBoxImage.Warning);

            if (result == Forms.MessageBoxResult.OK)
            {
                // start the transaction
                using (Transaction t = new Transaction(curDoc))
                {
                    t.Start("Delete Revisions");

                    // add a blank revision
                    Revision newRevision = Revision.Create(curDoc);

                    // get all the revisions in the project
                    IList<ElementId> revisions = Revision.GetAllRevisionIds(curDoc);

                    // remove the last revision from the list
                    revisions.RemoveAt(revisions.Count - 1);

                    // delete the remaining revisions
                    curDoc.Delete(revisions);

                    t.Commit();

                    string msgText2 = "All Revisions have been deleted.";
                    string msgTitle2 = "Complete";
                    Forms.MessageBoxButton msgButtons2 = Forms.MessageBoxButton.OK;

                    Forms.MessageBox.Show(msgText2, msgTitle2, msgButtons2, Forms.MessageBoxImage.Information);
                }

                return Result.Succeeded;
            }

            else
            {
                // exit the command
                return Result.Failed;
            }
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1_2";
            string buttonTitle = "Delete\rRevisions";
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
                    Properties.Resources.DeleteRevisions_32,
                    Properties.Resources.DeleteRevisions_16,
                    "Deletes revisions from project");

                return myButtonData1.Data;
            }
        }
    }
}

