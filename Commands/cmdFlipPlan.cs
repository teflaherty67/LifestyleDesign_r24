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

            // get Revit Command Id for Mirror Project
            RevitCommandId commandId = RevitCommandId.LookupPostableCommandId(PostableCommand.MirrorProject);

            // run the command using PostCommand
            uiapp.PostCommand(commandId);

            // run door swing reversal
            cmdReverseDoorSwings com_2 = new cmdReverseDoorSwings();
            com_2.Execute(commandData, ref message, elements);

            // run elevation rename
            cmdElevationRename com_3 = new cmdElevationRename();
            com_3.Execute(commandData, ref message, elements);

            // run sheet swap
            cmdElevationSheetSwap com_4 = new cmdElevationSheetSwap();
            com_4.Execute(commandData, ref message, elements);

            // run boundary shake
            cmdShakeAreaBoundary com_5 = new cmdShakeAreaBoundary();
            com_5.Execute(commandData, ref message, elements);

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";
            string? methodBase = MethodBase.GetCurrentMethod().DeclaringType?.FullName;

            if (methodBase == null)
            {
                throw new InvalidOperationException("MethodBase.GetCurrentMethod().DeclaringType?.FullName is null");
            }
            else
            {
                Classes.clsButtonDataClass myButtonData1 = new Classes.clsButtonDataClass(
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