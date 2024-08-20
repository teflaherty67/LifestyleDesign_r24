﻿using LifestyleDesign_r24.Classes;

namespace LifestyleDesign_r24
{
    [Transaction(TransactionMode.Manual)]
    public class cmdRevitMirror : IExternalCommand
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

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1_1";
            string buttonTitle = "Mirror\rProject";
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
                    "Mirrors project on specified axis");

                return myButtonData1.Data;
            }
        }
    }
}
