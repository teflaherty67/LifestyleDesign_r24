using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LifestyleDesign_r24
{
    [Transaction(TransactionMode.Manual)]
    public class cmdFlipPlan : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;            

            // run door swing reversal
            cmdReverseDoorSwings com_1 = new cmdReverseDoorSwings();
            com_1.Execute(commandData, ref message, elements);

            // run elevation rename
            cmdElevationRename com_2 = new cmdElevationRename();
            com_2.Execute(commandData, ref message, elements);

            // run sheet swap
            cmdElevationSheetSwap com_3 = new cmdElevationSheetSwap();
            com_3.Execute(commandData, ref message, elements);

            // run boundary shake
            cmdShakeAreaBoundary com_4 = new cmdShakeAreaBoundary();
            com_4.Execute(commandData, ref message, elements);

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand6_1";
            string buttonTitle = "Flip\rPlan";
            string methodBase = MethodBase.GetCurrentMethod().DeclaringType?.FullName;

            Bitmap curBmp32 = Classes.clsButtonData.ConvertBytetoBitmap(Properties.Resources.FlipPlan_32);
            Bitmap curBmp16 = Classes.clsButtonData.ConvertBytetoBitmap(Properties.Resources.FlipPlan_16);

            if (methodBase == null)
            {
                throw new InvalidOperationException("MethodBase.GetCurrentMethod().DeclaringType?.FullName is null");
            }
            else
            {
                Classes.clsButtonData myButtonData1 = new Classes.clsButtonData(
                    buttonInternalName,
                    buttonTitle,
                    methodBase,
                    curBmp32,
                    curBmp16,
                    "Completes the flipping process after the project is mirrored.");

                return myButtonData1.Data;
            }
        }
    }

}