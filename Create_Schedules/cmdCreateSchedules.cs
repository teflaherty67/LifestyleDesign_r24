﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LifestyleDesign_r24.Classes;
using LifestyleDesign_r24.Common;

namespace LifestyleDesign_r24
{
    [Transaction(TransactionMode.Manual)]
    public class cmdCreateSchedules : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document curDoc = uiapp.ActiveUIDocument.Document;

            // set some variables for paramter values
            string newFilter = "";

            if (GlobalVars.ElevDesignation == "A")
                newFilter = "1";
            else if (GlobalVars.ElevDesignation == "B")
                newFilter = "2";
            else if (GlobalVars.ElevDesignation == "C")
                newFilter = "3";
            else if (GlobalVars.ElevDesignation == "D")
                newFilter = "4";
            else if (GlobalVars.ElevDesignation == "S")
                newFilter = "5";
            else if (GlobalVars.ElevDesignation == "T")
                newFilter = "6";

            // open the form

            frmCreateSchedules curForm = new frmCreateSchedules()
            {
                Width = 420,
                Height = 400,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            curForm.ShowDialog();

            // get data from the form

            GlobalVars.ElevDesignation = curForm.GetComboBoxElevationSelectedItem();
            int floorNum = curForm.GetComboBoxFloorsSelectedItem();
            string typeFoundation = curForm.GetGroup1();
            string typeAttic = curForm.GetGroup2();


            bool chbIndexResult = curForm.GetCheckboxIndex();
            bool chbVeneerResult = curForm.GetCheckboxVeneer();
            bool chbFloorResult = curForm.GetCheckboxFloor();
            bool chbFrameResult = curForm.GetCheckboxFrame();
            bool chbAtticResult = curForm.GetCheckboxAttic();

            string levelWord = "";

            if (typeFoundation == "Basement" || typeFoundation == "Crawlspace")
            {
                levelWord = "Level";
            }
            else
            {
                levelWord = "Floor";
            }

            // start the transaction

            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Create Schedules");

                #region Sheet Index

                // check to see if the sheet index exists

                ViewSchedule schedIndex = Utils.GetScheduleByNameContains(curDoc, "Sheet Index - Elevation " + GlobalVars.ElevDesignation);

                if (chbIndexResult == true && schedIndex == null)
                {
                    Utils.DuplicateAndRenameSheetIndex(curDoc, newFilter);
                }

                #endregion

                #region Exterior Veneer Calculations

                ViewSchedule veneerIndex = Utils.GetScheduleByNameContains(curDoc, "Exterior Veneer Calculations - Elevation " + GlobalVars.ElevDesignation);

                if (chbVeneerResult == true && veneerIndex == null)
                {
                    Utils.DuplicateAndConfigureVeneerSchedule(curDoc);
                }

                #endregion

                #region Floor Areas                

                // check to see if the floor area scheme exists

                if (chbFloorResult == true)
                {
                    // set the variable for the floor Area Scheme name
                    AreaScheme floorAreaScheme = Utils.GetAreaSchemeByName(curDoc, GlobalVars.ElevDesignation + " Floor");

                    // set the variable for the color fill scheme
                    ColorFillScheme floorColorScheme = Utils.GetColorFillSchemeByName(curDoc, "Floor", floorAreaScheme);

                    if (floorAreaScheme == null || floorColorScheme == null)
                    {
                        // if null, warn the user & exit the command
                        TaskDialog tdFloorError = new TaskDialog("Error");
                        tdFloorError.MainIcon = Icon.TaskDialogIconWarning;
                        tdFloorError.Title = "Create Schedules";
                        tdFloorError.TitleAutoPrefix = false;
                        tdFloorError.MainContent = "Either the Area Scheme, or the Color Scheme, does not exist " +
                            "or is named incorrectly. Resolve the issue & try again.";
                        tdFloorError.CommonButtons = TaskDialogCommonButtons.Close;

                        TaskDialogResult tdFloorErrorRes = tdFloorError.Show();

                        return Result.Failed;
                    }

                    #region Floor Area Plans

                    // check if area plans exist
                    ViewPlan areaFloorView = Utils.GetAreaPlanByViewFamilyName(curDoc, GlobalVars.ElevDesignation + " Floor");

                    // if the floor area scheme exists, check to see if the floor area plans exist

                    if (floorAreaScheme != null)
                    {

                        // if not, create the area plans

                        if (areaFloorView == null)
                        {

                            List<Level> levelList = Utils.GetLevelByNameContains(curDoc, levelWord);

                            List<View> areaViews = new List<View>();

                            foreach (Level curlevel in levelList)
                            {
                                // get the category & set category Id
                                Category areaCat = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Areas);

                                // get the element Id of the current level
                                ElementId curLevelId = curlevel.Id;

                                // create & set variable for the view template
                                View vtFloorAreas = Utils.GetViewTemplateByName(curDoc, "10-Floor Area");

                                // create the area plan
                                ViewPlan areaFloor = ViewPlan.CreateAreaPlan(curDoc, floorAreaScheme.Id, curLevelId);
                                areaFloor.ViewTemplateId = vtFloorAreas.Id;
                                areaFloor.SetColorFillSchemeId(areaCat.Id, floorColorScheme.Id);

                                areaFloorView = areaFloor;

                                areaViews.Add(areaFloor);
                            }

                            foreach (ViewPlan curView in areaViews)
                            {
                                //set the color scheme                                   

                                // create insertion points
                                XYZ insStart = new XYZ(50, 0, 0);

                                UV insPoint = new UV(insStart.X, insStart.Y);
                                UV offset = new UV(0, 8);

                                XYZ tagInsert = new XYZ(50, 0, 0);
                                XYZ tagOffset = new XYZ(0, 8, 0);

                                if (curView.Name == "Lower Level")
                                {
                                    // add these areas
                                    List<clsAreaData> areasLower = new List<clsAreaData>()
                                        {
                                            new clsAreaData("13", "Living", "Total Covered", "A"),
                                            new clsAreaData("14", "Mechanical", "Total Covered", "K"),
                                            new clsAreaData("15", "Unfinished Basement", "Total Covered", "L"),
                                            new clsAreaData("16", "Option", "Options", "H")
                                        };
                                    foreach (var areaInfo in areasLower)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }

                                    // add the color fill legend
                                    Utils.AddColorLegend(curView, floorColorScheme);
                                }
                                else if (curView.Name == "Main Level" || curView.Name == "First Floor")
                                {
                                    // add these areas
                                    List<clsAreaData> areasMain = new List<clsAreaData>
                                        {
                                            new clsAreaData("1", "Living", "Total Covered", "A"),
                                            new clsAreaData("2", "Garage", "Total Covered", "B"),
                                            new clsAreaData("3", "Covered Patio", "Total Covered", "C"),
                                            new clsAreaData("4", "Covered Porch", "Total Covered", "D"),
                                            new clsAreaData("5", "Porte Cochere", "Total Covered", "E"),
                                            new clsAreaData("6", "Patio", "Total Uncovered", "F"),
                                            new clsAreaData("7", "Porch", "Total Uncovered", "G"),
                                            new clsAreaData("8", "Option", "Options", "H")
                                        };

                                    foreach (var areaInfo in areasMain)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }

                                    // add the color fill legend
                                    Utils.AddColorLegend(curView, floorColorScheme);
                                }
                                else
                                {
                                    // add these areas
                                    List<clsAreaData> areasUpper = new List<clsAreaData>
                                        {
                                            new clsAreaData("9", "Living", "Total Covered", "A"),
                                            new clsAreaData("10", "Covered Balcony", "Total Covered", "I"),
                                            new clsAreaData("11", "Balcony", "Total Uncovered", "J"),
                                            new clsAreaData("12", "Option", "Options", "H")
                                        };

                                    foreach (var areaInfo in areasUpper)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }

                                    // add the color fill legend
                                    Utils.AddColorLegend(curView, floorColorScheme);
                                }
                            }
                        }

                        #endregion

                        #region Floor Area Schedule

                        // if the floor area plans exist, create the schedule

                        // create & set variable for the view template
                        ViewSchedule vtFloorSched = Utils.GetViewScheduleTemplateByName(curDoc, "-Schedule-");

                        // create the new schedule
                        ViewSchedule newFloorSched = Utils.CreateAreaSchedule(curDoc, "Floor Areas - Elevation " + GlobalVars.ElevDesignation, floorAreaScheme);
                        newFloorSched.ViewTemplateId = vtFloorSched.Id;
                        Utils.SetParameterValue(newFloorSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

                        if (areaFloorView != null)
                        {
                            if (floorNum == 2 || floorNum == 3)
                            {
                                // get element Id of the fields to be used in the schedule
                                ElementId catFieldId = Utils.GetProjectParameterId(curDoc, "Area Category");
                                ElementId comFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                ElementId levelFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_LEVEL_ID);
                                ElementId nameFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_NAME);
                                ElementId areaFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_AREA);
                                ElementId numFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_NUMBER);

                                // create the fields and assign formatting properties
                                ScheduleField catField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, catFieldId);
                                catField.IsHidden = true;

                                ScheduleField comField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, comFieldId);
                                comField.IsHidden = true;

                                ScheduleField levelField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, levelFieldId);
                                levelField.IsHidden = false;
                                levelField.ColumnHeading = "Level";
                                levelField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                levelField.HorizontalAlignment = ScheduleHorizontalAlignment.Left;

                                ScheduleField nameField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                                nameField.IsHidden = true;

                                ScheduleField areaField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                                areaField.IsHidden = false;
                                areaField.ColumnHeading = "Area";
                                areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                                areaField.DisplayType = ScheduleFieldDisplayType.Totals;

                                FormatOptions formatOpts = new FormatOptions();
                                formatOpts.UseDefault = false;
                                formatOpts.SetUnitTypeId(UnitTypeId.SquareFeet);
                                formatOpts.SetSymbolTypeId(SymbolTypeId.Sf);

                                areaField.SetFormatOptions(formatOpts);

                                ScheduleField numField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, numFieldId);
                                numField.IsHidden = true;

                                // create the filters
                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.NotContains, "Options");
                                newFloorSched.Definition.AddFilter(catFilter);

                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.HasValue);
                                newFloorSched.Definition.AddFilter(areaFilter);

                                // set the sorting
                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                catSort.ShowFooter = true;
                                catSort.ShowFooterTitle = true;
                                catSort.ShowFooterCount = true;
                                catSort.ShowBlankLine = true;
                                newFloorSched.Definition.AddSortGroupField(catSort);

                                ScheduleSortGroupField comSort = new ScheduleSortGroupField(comField.FieldId, ScheduleSortOrder.Ascending);
                                newFloorSched.Definition.AddSortGroupField(comSort);

                                ScheduleSortGroupField nameSort = new ScheduleSortGroupField(nameField.FieldId, ScheduleSortOrder.Ascending);
                                nameSort.ShowHeader = true;
                                nameSort.ShowFooter = true;
                                newFloorSched.Definition.AddSortGroupField(nameSort);

                                ScheduleSortGroupField levelSort = new ScheduleSortGroupField(levelField.FieldId, ScheduleSortOrder.Ascending);
                                newFloorSched.Definition.AddSortGroupField(levelSort);

                                newFloorSched.Definition.IsItemized = false;
                            }
                            else
                            {
                                // get element Id of the fields to be used in the schedule
                                ElementId catFieldId = Utils.GetProjectParameterId(curDoc, "Area Category");
                                ElementId comFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                                ElementId nameFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_NAME);
                                ElementId areaFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_AREA);
                                ElementId numFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_NUMBER);

                                // create the fields and assign formatting properties
                                ScheduleField catField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, catFieldId);
                                catField.IsHidden = true;

                                ScheduleField comField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, comFieldId);
                                comField.IsHidden = true;

                                ScheduleField nameField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                                nameField.IsHidden = false;
                                nameField.ColumnHeading = "Name";
                                nameField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                nameField.HorizontalAlignment = ScheduleHorizontalAlignment.Left;

                                ScheduleField areaField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                                areaField.IsHidden = false;
                                areaField.ColumnHeading = "Area";
                                areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                                areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                                areaField.DisplayType = ScheduleFieldDisplayType.Totals;

                                FormatOptions formatOpts = new FormatOptions();
                                formatOpts.UseDefault = false;
                                formatOpts.SetUnitTypeId(UnitTypeId.SquareFeet);
                                formatOpts.SetSymbolTypeId(SymbolTypeId.Sf);

                                areaField.SetFormatOptions(formatOpts);

                                ScheduleField numField = newFloorSched.Definition.AddField(ScheduleFieldType.Instance, numFieldId);
                                numField.IsHidden = true;

                                // create the filters
                                ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.NotContains, "Options");
                                newFloorSched.Definition.AddFilter(catFilter);

                                ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.HasValue);
                                newFloorSched.Definition.AddFilter(areaFilter);

                                // set the sorting
                                ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                                catSort.ShowFooter = true;
                                catSort.ShowFooterTitle = true;
                                catSort.ShowFooterCount = true;
                                catSort.ShowBlankLine = true;
                                newFloorSched.Definition.AddSortGroupField(catSort);

                                ScheduleSortGroupField comSort = new ScheduleSortGroupField(comField.FieldId, ScheduleSortOrder.Ascending);
                                newFloorSched.Definition.AddSortGroupField(comSort);

                                newFloorSched.Definition.IsItemized = false;
                            }
                        }

                        #endregion
                    }
                }

                #endregion

                #region Frame Areas

                // check to see if the frame area scheme exists

                if (chbFrameResult == true)
                {
                    // set the variable for the frame Area Scheme name
                    AreaScheme frameAreaScheme = Utils.GetAreaSchemeByName(curDoc, GlobalVars.ElevDesignation + " Frame");

                    // set the variable for the color fill scheme
                    ColorFillScheme frameColorScheme = Utils.GetColorFillSchemeByName(curDoc, "Frame", frameAreaScheme);

                    if (frameAreaScheme == null || frameColorScheme == null)
                    {
                        // if null, warn the user & exit the command
                        TaskDialog tdFrameError = new TaskDialog("Error");
                        tdFrameError.MainIcon = Icon.TaskDialogIconWarning;
                        tdFrameError.Title = "Create Schedules";
                        tdFrameError.TitleAutoPrefix = false;
                        tdFrameError.MainContent = "Either the Area Scheme, or the Color Scheme, does not exist " +
                            "or is named incorrectly. Resolve the issue & try again.";
                        tdFrameError.CommonButtons = TaskDialogCommonButtons.Close;

                        TaskDialogResult tdFrameErrorRes = tdFrameError.Show();
                    }

                    #region Frame Area Plans

                    // check to see if the frame area plans exist
                    ViewPlan areaFrameView = Utils.GetAreaPlanByViewFamilyName(curDoc, GlobalVars.ElevDesignation + " Frame");

                    // if the frame area scheme exists, check to see if the frame area plans exist
                    if (frameAreaScheme != null)
                    {
                        // if not, create the area plans
                        if (areaFrameView == null)
                        {
                            List<Level> levelList = Utils.GetLevelByNameContains(curDoc, levelWord);

                            List<View> areaViews = new List<View>();

                            // loop through the levels and create the views
                            foreach (Level curLevel in levelList)
                            {
                                // get the category & set the category Id
                                Category areaCat = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Areas);

                                // get the element Id of the current level
                                ElementId curLevelId = curLevel.Id;

                                // create & set variable for the view template
                                View vtFrameAreas = Utils.GetViewTemplateByName(curDoc, "11-Frame Area");

                                // create the area plan
                                ViewPlan areaFrame = ViewPlan.CreateAreaPlan(curDoc, frameAreaScheme.Id, curLevelId);
                                areaFrame.ViewTemplateId = vtFrameAreas.Id;
                                areaFrame.SetColorFillSchemeId(areaCat.Id, frameColorScheme.Id);

                                areaFrameView = areaFrame;

                                areaViews.Add(areaFrame);
                            }

                            foreach (ViewPlan curView in areaViews)
                            {
                                // create insertion points
                                XYZ insStart = new XYZ(50, 0, 0);

                                UV insPoint = new UV(insStart.X, insStart.Y);
                                UV offset = new UV(0, 8);

                                XYZ tagInsert = new XYZ(50, 0, 0);
                                XYZ tagOffset = new XYZ(0, 8, 0);

                                if (curView.Name == "Lower Level")
                                {
                                    // add these areas
                                    List<clsAreaData> areasLower = new List<clsAreaData>()
                                    {
                                        new clsAreaData("5", "Standard", "", ""),
                                        new clsAreaData("6", "Option", "Options", "")
                                    };
                                    foreach (var areaInfo in areasLower)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }

                                    // add the color fill legend
                                    Utils.AddColorLegend(curView, frameColorScheme);
                                }
                                else if (curView.Name == "Main Level" || curView.Name == "First Floor")
                                {
                                    // add these areas
                                    List<clsAreaData> areasMain = new List<clsAreaData>()
                                    {
                                        new clsAreaData("1", "Standard", "", ""),
                                        new clsAreaData("2", "Option", "Options", "")
                                    };
                                    foreach (var areaInfo in areasMain)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }

                                    // add the color fill legend
                                    Utils.AddColorLegend(curView, frameColorScheme);
                                }
                                else
                                {
                                    // add these areas
                                    List<clsAreaData> areasUpper = new List<clsAreaData>()
                                    {
                                        new clsAreaData("3", "Standard", "", ""),
                                        new clsAreaData("4", "Option", "Options", "")
                                    };
                                    foreach (var areaInfo in areasUpper)
                                    {
                                        Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                                    }

                                    // add the color fill legend
                                    Utils.AddColorLegend(curView, frameColorScheme);
                                }
                            }
                        }
                    }

                    #endregion

                    #region Frame Area Schedule

                    // if the frame area plans, exist create the schedule

                    // create & set variable for the view template
                    ViewSchedule vtFrameSched = Utils.GetViewScheduleTemplateByName(curDoc, "-Frame Areas-");

                    // create the new schedule
                    ViewSchedule newFrameSched = Utils.CreateAreaSchedule(curDoc, "Frame Areas - Elevation " + GlobalVars.ElevDesignation, frameAreaScheme);
                    newFrameSched.ViewTemplateId = vtFrameSched.Id;
                    Utils.SetParameterValue(newFrameSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

                    if (areaFrameView != null)
                    {
                        if (floorNum == 2 || floorNum == 3)
                        {
                            // get the element Id fo the fields to be used in the schedule
                            ElementId catFieldId = Utils.GetProjectParameterId(curDoc, "Area Category");
                            ElementId nameFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_NAME);
                            ElementId levelFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_LEVEL_ID);
                            ElementId areaFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_AREA);

                            // create the fields & set the formatting properties
                            ScheduleField catField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, catFieldId);
                            catField.IsHidden = true;

                            ScheduleField nameField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                            nameField.IsHidden = true;

                            ScheduleField levelField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, levelFieldId);
                            levelField.IsHidden = false;
                            levelField.ColumnHeading = "Level";
                            levelField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                            levelField.HorizontalAlignment = ScheduleHorizontalAlignment.Left;

                            ScheduleField areaField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                            areaField.IsHidden = false;
                            areaField.ColumnHeading = "Area";
                            areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                            areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                            areaField.DisplayType = ScheduleFieldDisplayType.Totals;

                            FormatOptions formatOpts = new FormatOptions();
                            formatOpts.UseDefault = false;
                            formatOpts.SetUnitTypeId(UnitTypeId.SquareFeet);
                            formatOpts.SetSymbolTypeId(SymbolTypeId.Sf);

                            areaField.SetFormatOptions(formatOpts);

                            // create the filters
                            ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.NotContains, "Options");
                            newFrameSched.Definition.AddFilter(catFilter);

                            ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.GreaterThan, 0.0);

                            // set the sorting
                            ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                            catSort.ShowHeader = true;
                            catSort.ShowBlankLine = true;
                            newFrameSched.Definition.AddSortGroupField(catSort);

                            ScheduleSortGroupField nameSort = new ScheduleSortGroupField(nameField.FieldId, ScheduleSortOrder.Ascending);
                            nameSort.ShowHeader = true;
                            nameSort.ShowFooter = true;
                            newFrameSched.Definition.AddSortGroupField(nameSort);

                            ScheduleSortGroupField levelSort = new ScheduleSortGroupField(levelField.FieldId, ScheduleSortOrder.Ascending);
                            newFrameSched.Definition.AddSortGroupField(levelSort);

                            newFrameSched.Definition.IsItemized = false;
                        }
                        else
                        {
                            // get the element Id fo the fields to be used in the schedule
                            ElementId catFieldId = Utils.GetProjectParameterId(curDoc, "Area Category");
                            ElementId nameFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_NAME);
                            ElementId areaFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_AREA);

                            // create the fields & set the formatting properties
                            ScheduleField catField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, catFieldId);
                            catField.IsHidden = true;

                            ScheduleField nameField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                            nameField.IsHidden = false;
                            nameField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                            nameField.HorizontalAlignment = ScheduleHorizontalAlignment.Left;

                            ScheduleField areaField = newFrameSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                            areaField.IsHidden = false;
                            areaField.ColumnHeading = "Area";
                            areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                            areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Right;
                            areaField.DisplayType = ScheduleFieldDisplayType.Totals;

                            FormatOptions formatOpts = new FormatOptions();
                            formatOpts.UseDefault = false;
                            formatOpts.SetUnitTypeId(UnitTypeId.SquareFeet);
                            formatOpts.SetSymbolTypeId(SymbolTypeId.Sf);

                            areaField.SetFormatOptions(formatOpts);

                            // create the filters
                            ScheduleFilter catFilter = new ScheduleFilter(catField.FieldId, ScheduleFilterType.NotContains, "Options");
                            newFrameSched.Definition.AddFilter(catFilter);

                            ScheduleFilter areaFilter = new ScheduleFilter(areaField.FieldId, ScheduleFilterType.GreaterThan, 0.0);

                            // set the sorting
                            ScheduleSortGroupField catSort = new ScheduleSortGroupField(catField.FieldId, ScheduleSortOrder.Ascending);
                            catSort.ShowHeader = true;
                            catSort.ShowBlankLine = true;

                            newFrameSched.Definition.IsItemized = false;
                        }
                    }
                    #endregion
                }
                #endregion

                #region Attic Areas

                // check to see if the attic area scheme exists

                if (chbAtticResult == true)
                {
                    // set the variable for the attic area scheme name
                    AreaScheme atticAreaScheme = Utils.GetAreaSchemeByName(curDoc, GlobalVars.ElevDesignation + " Roof Ventilation");

                    // set the variable for the color fill scheme
                    ColorFillScheme atticColorScheme = Utils.GetColorFillSchemeByName(curDoc, "Attic", atticAreaScheme);

                    if (atticAreaScheme == null || atticColorScheme == null)
                    {
                        // if null, warn the user & exit the command
                        TaskDialog tdAtticError = new TaskDialog("Error");
                        tdAtticError.MainIcon = Icon.TaskDialogIconWarning;
                        tdAtticError.Title = "Create Schedules";
                        tdAtticError.TitleAutoPrefix = false;
                        tdAtticError.MainContent = "Either the attic Area Scheme, or the attic Color Scheme, does not exist " +
                            "or is named incorrectly. Resolve the issue & try again.";
                        tdAtticError.CommonButtons = TaskDialogCommonButtons.Close;

                        TaskDialogResult tdAtticErrorRes = tdAtticError.Show();

                        return Result.Failed;
                    }

                    #region Attic Area Plans

                    // check if the attic area plans exists
                    ViewPlan areaAtticView = Utils.GetAreaPlanByViewFamilyName(curDoc, GlobalVars.ElevDesignation + " Roof Ventilation");

                    // create a variable for the attic level
                    string atticLevel = "";

                    // if the attic area scheme exists, check if the attic area plans exist
                    if (atticAreaScheme != null)
                    {
                        // if not, create the area plans
                        if (areaAtticView == null)
                        {
                            if (typeFoundation == "Basement" || typeFoundation == "Crawlspace" && floorNum == 3)
                            {
                                atticLevel = "Upper Level";
                            }
                            else if (typeFoundation == "Basement" || typeFoundation == "Crawlspace" && floorNum == 2)
                            {
                                atticLevel = "Main Level";
                            }
                            else if (typeFoundation == "Slab" && floorNum == 3)
                            {
                                atticLevel = "Third Floor";
                            }
                            else if (typeFoundation == "Slab" && floorNum == 2)
                            {
                                atticLevel = "Second Floor";
                            }
                            else
                            {
                                atticLevel = "First Floor";
                            }

                            // get the category
                            Category areaCat = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Areas);

                            // get the level for attic area plan
                            Level levelAttic = Utils.GetLevelByName(curDoc, atticLevel);

                            // create & set variable for the view template
                            View vtAtticAreas = Utils.GetViewTemplateByName(curDoc, "12-Attic Area");

                            // create the area plan
                            ViewPlan areaAttic = ViewPlan.CreateAreaPlan(curDoc, atticAreaScheme.Id, levelAttic.Id);
                            areaAttic.Name = "Roof";
                            areaAttic.ViewTemplateId = vtAtticAreas.Id;
                            areaAttic.SetColorFillSchemeId(areaCat.Id, atticColorScheme.Id);

                            ViewPlan curView = areaAttic;

                            // create insertion points                            
                            XYZ insStart = new XYZ(50, 0, 0);

                            UV insPoint = new UV(insStart.X, insStart.Y);
                            UV offset = new UV(0, 8);

                            XYZ tagInsert = new XYZ(50, 0, 0);
                            XYZ tagOffset = new XYZ(0, 8, 0);

                            // add these areas
                            List<clsAreaData> areasAttic = new List<clsAreaData>()
                            {
                                new clsAreaData("1", "Attic 1", 0),
                                new clsAreaData("2", "Attic 2", 0)
                            };
                            foreach (var areaInfo in areasAttic)
                            {
                                Utils.CreateFloorAreaWithTag(curDoc, curView, ref insPoint, ref tagInsert, areaInfo);
                            }

                            // add the color fill legend
                            Utils.AddColorLegend(curView, atticColorScheme);
                        }
                    }

                    #endregion

                    #region Roof Ventilation Calculations

                    // if the attic plans exist, create the ventilation calculations schedule

                    // create & set variable for the view template
                    ViewSchedule vtAtticSched = Utils.GetViewScheduleTemplateByName(curDoc, "-Schedule-");

                    // create the new schedule
                    ViewSchedule newAtticSched = Utils.CreateAreaSchedule(curDoc, "Roof Ventilation Calculations - Elevation " + GlobalVars.ElevDesignation, atticAreaScheme);
                    newAtticSched.ViewTemplateId = vtAtticSched.Id;
                    Utils.SetParameterValue(newAtticSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

                    if (typeAttic == "Multi-Space")
                    {
                        // get the element Id of the fields to be used in the schedule
                        ElementId nameFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_NAME);
                        ElementId areaFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_AREA);
                        ElementId ratioFieldId = Utils.GetProjectParameterId(curDoc, "150 Ratio");

                        // create the fields & set the formatting properties
                        ScheduleField nameField = newAtticSched.Definition.AddField(ScheduleFieldType.Instance, nameFieldId);
                        nameField.IsHidden = false;
                        nameField.ColumnHeading = "Named Attic Space";
                        nameField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                        nameField.HorizontalAlignment = ScheduleHorizontalAlignment.Center;

                        ScheduleField areaField = newAtticSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                        areaField.IsHidden = false;
                        areaField.ColumnHeading = "Attic Area";
                        areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                        areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                        areaField.DisplayType = ScheduleFieldDisplayType.Totals;

                        FormatOptions formatOpts = new FormatOptions();
                        formatOpts.UseDefault = false;
                        formatOpts.SetUnitTypeId(UnitTypeId.SquareFeet);
                        formatOpts.SetSymbolTypeId(SymbolTypeId.Sf);

                        areaField.SetFormatOptions(formatOpts);

                        ScheduleField ratioField = newAtticSched.Definition.AddField(ScheduleFieldType.Instance, ratioFieldId);
                        ratioField.IsHidden = true;
                    }
                    else
                    {
                        // get the element Id of the fields to be used in the schedule
                        ElementId areaFieldId = Utils.GetBuiltInParameterId(curDoc, BuiltInCategory.OST_Areas, BuiltInParameter.ROOM_AREA);
                        ElementId ratioFieldId = Utils.GetProjectParameterId(curDoc, "150 Ratio");

                        // create the fields & set the formatting properties                        
                        ScheduleField areaField = newAtticSched.Definition.AddField(ScheduleFieldType.Instance, areaFieldId);
                        areaField.IsHidden = false;
                        areaField.ColumnHeading = "Attic Area";
                        areaField.HeadingOrientation = ScheduleHeadingOrientation.Horizontal;
                        areaField.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                        areaField.DisplayType = ScheduleFieldDisplayType.Totals;

                        FormatOptions formatOpts = new FormatOptions();
                        formatOpts.UseDefault = false;
                        formatOpts.SetUnitTypeId(UnitTypeId.SquareFeet);
                        formatOpts.SetSymbolTypeId(SymbolTypeId.Sf);

                        areaField.SetFormatOptions(formatOpts);

                        ScheduleField ratioField = newAtticSched.Definition.AddField(ScheduleFieldType.Instance, ratioFieldId);
                        ratioField.IsHidden = true;
                    }

                    #endregion

                    #region Roof Ventilation Equipment

                    // search for equipement schedule
                    ViewSchedule equipmentSched = Utils.GetScheduleByNameContains(curDoc, "Roof Ventilation Equipment - Elevation " + GlobalVars.ElevDesignation);

                    // if not found, create it
                    if (chbAtticResult == true && equipmentSched == null)
                    {
                        Utils.DuplicateAndConfigureEquipmentSchedule(curDoc);
                    }

                    #endregion
                }

                #endregion

                t.Commit();
            }

            // alert the user
            // if null, warn the user & exit the command
            TaskDialog tdSchedSuccess = new TaskDialog("Complete");
            tdSchedSuccess.MainIcon = Icon.TaskDialogIconInformation;
            tdSchedSuccess.Title = "Create Schedules";
            tdSchedSuccess.TitleAutoPrefix = false;
            tdSchedSuccess.MainContent = "The specified schedules have been created. However, the API is not capable of setting the Design Option or adding calculated paramters. " +
                "The Design Option for the Exterior Veneer Calculations and the Roof Ventilation Equipment schedules will need to be set manually. Also," +
                "the calculated parameters, Ratio & Required Free Ventilation Area, for the Roof Ventilation Calculaitons schedule will need to be added.";
            tdSchedSuccess.CommonButtons = TaskDialogCommonButtons.Close;

            TaskDialogResult tdSchedSuccessRes = tdSchedSuccess.Show();

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
                clsButtonData myButtonData1 = new clsButtonData(
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
