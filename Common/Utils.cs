using LifestyleDesign_r24.Classes;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;

namespace LifestyleDesign_r24.Common
{
    internal static class Utils
    {
        #region Area Plans

        internal static void CreateFloorAreaWithTag(Document curDoc, ViewPlan areaPlan, ref UV insPoint, ref XYZ tagInsert, clsAreaData areaInfo)
        {
            Area curArea = curDoc.Create.NewArea(areaPlan, insPoint);
            curArea.Number = areaInfo.Number;
            curArea.Name = areaInfo.Name;
            curArea.LookupParameter("Area Category").Set(areaInfo.Category);
            curArea.LookupParameter("Comments").Set(areaInfo.Comments);

            AreaTag tag = curDoc.Create.NewAreaTag(areaPlan, curArea, insPoint);
            tag.TagHeadPosition = tagInsert;
            tag.HasLeader = false;

            UV offset = new UV(0, 8);
            insPoint = insPoint.Subtract(offset);

            XYZ tagOffset = new XYZ(0, 8, 0);
            tagInsert = tagInsert.Subtract(tagOffset);

            if (areaInfo.Ratio != 99)
            {
                curArea.LookupParameter("150 Ratio").Set(areaInfo.Ratio);
            }
        }

        internal static ViewPlan GetAreaPlanByViewFamilyName(Document doc, string vftName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewPlan));

            foreach (ViewPlan curViewPlan in collector)
            {
                if (curViewPlan.ViewType == ViewType.AreaPlan)
                {
                    ViewFamilyType curVFT = doc.GetElement(curViewPlan.GetTypeId()) as ViewFamilyType;

                    if (curVFT.Name == vftName)
                        return curViewPlan;
                }
            }

            return null;
        }

        #endregion        

        #region Area Scheme

        internal static AreaScheme GetAreaSchemeByName(Document doc, string schemeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(AreaScheme));

            foreach (AreaScheme areaScheme in collector)
            {
                if (areaScheme.Name == schemeName)
                {
                    return areaScheme;
                }
            }

            return null;
        }

        #endregion

        #region Color Fill Scheme

        internal static ColorFillScheme GetColorFillSchemeByName(Document curDoc, string schemeName, AreaScheme areaScheme)
        {
            try
            {
                ColorFillScheme colorfill = new FilteredElementCollector(curDoc)
               .OfCategory(BuiltInCategory.OST_ColorFillSchema)
               .Cast<ColorFillScheme>()
               .Where(x => x.Name.Equals(schemeName) && x.AreaSchemeId.Equals(areaScheme.Id))
               .First();

                return colorfill;
            }
            catch
            {
                return null;
            }
        }

        internal static void AddColorLegend(View view, ColorFillScheme scheme)
        {
            ElementId areaCatId = new ElementId(BuiltInCategory.OST_Areas);
            ElementId curLegendId = view.GetColorFillSchemeId(areaCatId);

            if (curLegendId == ElementId.InvalidElementId)
                view.SetColorFillSchemeId(areaCatId, scheme.Id);

            ColorFillLegend.Create(view.Document, view.Id, areaCatId, XYZ.Zero);
        }

        #endregion

        #region Levels

        public static List<Level> GetAllLevels(Document doc)
        {
            FilteredElementCollector colLevels = new FilteredElementCollector(doc);
            colLevels.OfCategory(BuiltInCategory.OST_Levels);

            List<Level> levels = new List<Level>();
            foreach (Element x in colLevels.ToElements())
            {
                if (x.GetType() == typeof(Level))
                {
                    levels.Add((Level)x);
                }
            }

            return levels;           
        }

        internal static List<Level> GetLevelByNameContains(Document doc, string levelWord)
        {
            List<Level> levels = GetAllLevels(doc);

            List<Level> returnList = new List<Level>();

            foreach (Level curLevel in levels)
            {
                if (curLevel.Name.Contains(levelWord))
                    returnList.Add(curLevel);
            }

            return returnList;
        }

        internal static Level GetLevelByName(Document doc, string levelName)
        {
            List<Level> levels = GetAllLevels(doc);

            foreach (Level curLevel in levels)
            {
                Debug.Print(curLevel.Name);

                if (curLevel.Name.Equals(levelName))
                    return curLevel;
            }

            return null;
        }

        #endregion

        #region Parameters

        internal static string GetParameterValueByName(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
                try
                {
                    Parameter param = paramList[0];
                    string paramValue = param.AsValueString();
                    return paramValue;
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    return null;
                }

            return "";
        }        

        internal static Parameter GetParameterByName(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName)
                    return curParam;
            }

            return null;
        }

        internal static Parameter GetParameterByNameAndWritable(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName && curParam.IsReadOnly == false)
                    return curParam;
            }

            return null;
        }

        internal static ElementId GetProjectParameterId(Document doc, string name)
        {
            ParameterElement pElem = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterElement))
                .Cast<ParameterElement>()
                .Where(e => e.Name.Equals(name))
                .FirstOrDefault();

            return pElem?.Id;
        }

        internal static ElementId GetBuiltInParameterId(Document doc, BuiltInCategory cat, BuiltInParameter bip)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(cat);

            Parameter curParam = collector.FirstElement().get_Parameter(bip);

            return curParam?.Id;
        }

        internal static string SetParameterByNameAndWritable(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByNameAndWritable(curElem, paramName);

            curParam.Set(value);
            return curParam.ToString();
        }

        internal static void SetParameterByName(Element element, string paramName, string value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
            {
                Parameter param = paramList[0];

                param.Set(value);
            }
        }

        internal static void SetParameterByName(Element element, string paramName, int value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
            {
                Parameter param = paramList[0];

                param.Set(value);
            }
        }

        internal static bool SetParameterValue(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);

            if (curParam != null)
            {
                curParam.Set(value);
                return true;
            }

            return false;
        }

        #endregion

        #region Ribbon

        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel currentPanel = GetRibbonPanelByName(app, tabName, panelName);

            if (currentPanel == null)
                currentPanel = app.CreateRibbonPanel(tabName, panelName);

            return currentPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        internal static BitmapImage BitmapToImageSource(Bitmap bm)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }

        #endregion

        #region Schedules

        internal static ViewSchedule CreateAreaSchedule(Document doc, string schedName, AreaScheme curAreaScheme)
        {
            ElementId catId = new ElementId(BuiltInCategory.OST_Areas);
            ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId, curAreaScheme.Id);
            newSchedule.Name = schedName;

            return newSchedule;
        }

        internal static List<ViewSchedule> GetAllSchedulesByElevation(Document doc, string newElev)
        {
            List<ViewSchedule> m_scheduleList = GetAllSchedules(doc);

            List<ViewSchedule> m_returnList = new List<ViewSchedule>();

            foreach (ViewSchedule curVS in m_scheduleList)
            {
                if (curVS.Name.EndsWith(newElev))
                {
                    m_returnList.Add(curVS);
                }
            }

            return m_returnList;
        }

        internal static List<ViewSchedule> GetAllSchedules(Document doc)
        {
            List<ViewSchedule> schedList = new List<ViewSchedule>();

            FilteredElementCollector curCollector = new FilteredElementCollector(doc);
            curCollector.OfClass(typeof(ViewSchedule));

            //loop through views and check if schedule - if so then put into schedule list
            foreach (ViewSchedule curView in curCollector)
            {
                if (curView.ViewType == ViewType.Schedule)
                {
                    schedList.Add((ViewSchedule)curView);
                }
            }

            return schedList;
        }

        internal static ViewSchedule GetScheduleByName(Document doc, string v)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewSchedule));

            foreach (ViewSchedule curSchedule in collector)
            {
                if (curSchedule.Name == v)
                    return curSchedule;
            }

            return null;
        }

        internal static List<ViewSchedule> GetAllScheduleByNameContains(Document doc, string schedName)
        {
            List<ViewSchedule> m_scheduleList = GetAllSchedules(doc);

            List<ViewSchedule> m_returnList = new List<ViewSchedule>();

            foreach (ViewSchedule curSchedule in m_scheduleList)
            {
                if (curSchedule.Name.Contains(schedName))
                    m_returnList.Add(curSchedule);
            }

            return m_returnList;
        }

        private static List<ViewSchedule> GetAllViewScheduleTemplates(Document curDoc)
        {
            List<ViewSchedule> returnList = new List<ViewSchedule>();
            List<ViewSchedule> viewList = GetAllSchedules(curDoc);

            //loop through views and check if is view template
            foreach (ViewSchedule v in viewList)
            {
                if (v.IsTemplate == true)
                {
                    //add view template to list
                    returnList.Add(v);
                }
            }

            return returnList;
        }

        public static ViewSchedule GetViewScheduleTemplateByName(Document curDoc, string viewSchedTemplateName)
        {
            List<ViewSchedule> viewSchedTemplateList = GetAllViewScheduleTemplates(curDoc);

            foreach (ViewSchedule v in viewSchedTemplateList)
            {
                if (v.Name == viewSchedTemplateName)
                {
                    return v;
                }
            }

            return null;
        }

        internal static List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByNameAndView(Document doc, string elevName, View activeView)
        {
            List<ScheduleSheetInstance> ssiList = GetAllScheduleSheetInstancesByView(doc, activeView);

            List<ScheduleSheetInstance> returnList = new List<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance curInstance in ssiList)
            {
                if (curInstance.Name.Contains(elevName))
                    returnList.Add(curInstance);
            }

            return returnList;
        }

        internal static List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByView(Document doc, View activeView)
        {
            FilteredElementCollector colSSI = new FilteredElementCollector(doc, activeView.Id);
            colSSI.OfClass(typeof(ScheduleSheetInstance));

            List<ScheduleSheetInstance> returnList = new List<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance curInstance in colSSI)
            {
                returnList.Add(curInstance);
            }

            return returnList;
        }

        internal static List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByName(Document doc, string elevName)
        {
            List<ScheduleSheetInstance> ssiList = GetAllScheduleSheetInstances(doc);

            List<ScheduleSheetInstance> returnList = new List<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance curInstance in ssiList)
            {
                if (curInstance.Name.Contains(elevName))
                    returnList.Add(curInstance);
            }

            return returnList;
        }

        internal static List<ScheduleSheetInstance> GetAllScheduleSheetInstances(Document doc)
        {
            FilteredElementCollector colSSI = new FilteredElementCollector(doc);
            colSSI.OfClass(typeof(ScheduleSheetInstance));

            List<ScheduleSheetInstance> returnList = new List<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance curInstance in colSSI)
            {
                returnList.Add(curInstance);
            }

            return returnList;
        }

        internal static ViewSchedule GetScheduleByNameContains(Document doc, string scheduleString)
        {
            List<ViewSchedule> m_scheduleList = GetAllSchedules(doc);

            foreach (ViewSchedule curSchedule in m_scheduleList)
            {
                if (curSchedule.Name.Contains(scheduleString))
                    return curSchedule;
            }

            return null;
        }

        internal static void DuplicateAndRenameSheetIndex(Document curDoc, string newFilter)
        {
            // duplicate the first schedule found with "Sheet Index" in the name
            List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Sheet Index");
            ViewSchedule dupSched = listSched.FirstOrDefault();

            if (dupSched == null)
            {
                // call another method to create one

                return; // no schedule to duplicate
            }

            ViewSchedule indexSched = curDoc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;

            // rename the duplicated schedule to the new elevation
            string originalName = indexSched.Name;
            string[] schedTitle = originalName.Split('-');

            string curTitle = schedTitle[0];

            indexSched.Name = schedTitle[0] + "- Elevation " + GlobalVars.ElevDesignation;

            Utils.SetParameterValue(indexSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

            // update the filter value to the new elevation code filter
            ScheduleFilter codeFilter = indexSched.Definition.GetFilter(0);

            if (codeFilter.IsStringValue)
            {
                codeFilter.SetValue(newFilter);
                indexSched.Definition.SetFilter(0, codeFilter);
            }
        }

        internal static void DuplicateAndConfigureVeneerSchedule(Document curDoc)
        {
            // duplicate the first schedule with "Exterior Venner Calculations" in the name
            List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Exterior Veneer Calculations");
            ViewSchedule dupSched = listSched.FirstOrDefault();

            if (dupSched == null)
            {
                // call another method to create one

                return; // no schedule to duplicate
            }

            // duplicate the schedule
            ViewSchedule veneerSched = curDoc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;

            // rename the duplicated schedule to the new elevation
            string originalName = veneerSched.Name;
            string[] schedTitle = originalName.Split('-');

            veneerSched.Name = schedTitle[0] + "- Elevation " + GlobalVars.ElevDesignation;

            Utils.SetParameterValue(veneerSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

            //// set the design option to the specified elevation designation
            //DesignOption curOption = Utils.getDesignOptionByName(curDoc, "Elevation : " + Globals.ElevDesignation);

            //Parameter doParam = veneerSched.get_Parameter(BuiltInParameter.VIEWER_OPTION_VISIBILITY);

            //doParam.Set(curOption.Id); //??? the code is getting the right option, but it's not changing anything in the model
        }

        internal static void DuplicateAndConfigureEquipmentSchedule(Document curDoc)
        {
            // duplicate the first schedule with "Roof Ventilation Equipment" in the name
            List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Roof Ventilation Equipment");
            ViewSchedule dupSched = listSched.FirstOrDefault();

            if (dupSched == null)
            {
                // call another method to create one

                return; // no schedule to duplicate
            }

            // duplicate the schedule
            ViewSchedule equipmentSched = curDoc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;

            // rename the duplicated schedule to the new elevation
            string originalName = equipmentSched.Name;
            string[] schedTitle = originalName.Split('-');

            equipmentSched.Name = schedTitle[0] + "- Elevation " + GlobalVars.ElevDesignation;

            Utils.SetParameterValue(equipmentSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

            //// set the design option to the specified elevation designation
            //DesignOption curOption = Utils.getDesignOptionByName(curDoc, "Elevation : " + Globals.ElevDesignation);

            //Parameter doParam = veneerSched.get_Parameter(BuiltInParameter.VIEWER_OPTION_VISIBILITY);

            //doParam.Set(curOption.Id); //??? the code is getting the right option, but it's not changing anything in the model
        }

        #endregion

        #region Sheets

        internal static List<ViewSheet> GetAllSheets(Document curDoc)
        {
            //get all sheets
            FilteredElementCollector m_colSheets = new FilteredElementCollector(curDoc);
            m_colSheets.OfCategory(BuiltInCategory.OST_Sheets);

            List<ViewSheet> m_returnSheets = new List<ViewSheet>();
            foreach (ViewSheet curSheet in m_colSheets.ToElements())
            {
                m_returnSheets.Add(curSheet);
            }

            return m_returnSheets;
        }

        internal static ViewSheet GetSheetByElevationAndNameContains(Document curDoc, string newElev, string sheetName)
        {
            List<ViewSheet> sheetList = GetAllSheets(curDoc);

            foreach (ViewSheet curVS in sheetList)
            {
                if (curVS.SheetNumber.Contains(newElev.ToLower()) && curVS.Name.Contains(sheetName))
                {
                    return curVS;
                }
            }

            return null;
        }

        internal static List<ViewSheet> GetAllSheetsByElevation(Document curDoc, string elevDesignation)
        {
            //get all sheets
            List<ViewSheet> m_sheetList = GetAllSheets(curDoc);

            List<ViewSheet> m_returnSheets = new List<ViewSheet>();

            foreach (ViewSheet curSheet in m_sheetList)
            {
                if (curSheet.SheetNumber.Contains(elevDesignation))
                    m_returnSheets.Add(curSheet);
            }

            return m_returnSheets;
        }

        public static List<ViewSheet> GetSheetsByNumber(Document curDoc, string sheetNumber)
        {
            List<ViewSheet> returnSheets = new List<ViewSheet>();

            //get all sheets
            List<ViewSheet> curSheets = GetAllSheets(curDoc);

            //loop through sheets and check sheet name
            foreach (ViewSheet curSheet in curSheets)
            {
                if (curSheet.SheetNumber.Contains(sheetNumber))
                {
                    returnSheets.Add(curSheet);
                }
            }

            return returnSheets;
        }

        #endregion

        #region Strings

        internal static string GetLastCharacterInString(string grpName, string curElev, string newElev)
        {
            char lastChar = grpName[grpName.Length - 1];

            char firstChar = grpName[0];

            string grpLastChar = lastChar.ToString();

            string grpFirstChar = firstChar.ToString();

            if (grpLastChar == curElev)
            {
                return "Elevation " + newElev;
            }
            else if (grpFirstChar == curElev)
            {
                return newElev + grpName.Substring(1);
            }
            else
            {
                return grpName;
            }
        }

        internal static string GetStringBetweenCharacters(string input, string charFrom, string charTo)
        {
            //string cleanInput = CleanSheetNumber(input);

            int posFrom = input.IndexOf(charFrom);
            if (posFrom != -1) //if found char
            {
                int posTo = input.IndexOf(charTo, posFrom + 1);
                if (posTo != -1) //if found char
                {
                    return input.Substring(posFrom + 1, posTo - posFrom - 1);
                }
            }

            return string.Empty;
        }

        internal static string CleanSheetNumber(string sheetNumber)
        {
            Regex regex = new Regex(@"[^a-zA-Z0-9\s]", (RegexOptions)0);
            string replaceText = regex.Replace(sheetNumber, "");

            return replaceText;
        }

        #endregion

        #region Titleblocks

        internal static FamilySymbol GetTitleBlockByNameContains(Document curDoc, string tBlockName)
        {
            FilteredElementCollector m_tBlocks = new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsElementType();

            foreach (FamilySymbol curTBlock in m_tBlocks)
            {
                if (curTBlock.Name.Contains(tBlockName))
                    return curTBlock;
            }

            return null;
        }

        #endregion

        #region Views

        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector m_colviews = new FilteredElementCollector(curDoc);
            m_colviews.OfCategory(BuiltInCategory.OST_Views);

            List<View> m_returnViews = new List<View>();
            foreach (View curView in m_colviews.ToElements())
            {
                m_returnViews.Add(curView);
            }

            return m_returnViews;
        }

        public static List<View> GetAllElevationViews(Document doc)
        {
            List<View> returnList = new List<View>();

            FilteredElementCollector colViews = new FilteredElementCollector(doc);
            colViews.OfClass(typeof(View));

            // loop through views and check for elevation views
            foreach (View x in colViews)
            {
                if (x.GetType() == typeof(ViewSection))
                {
                    if (x.IsTemplate == false)
                    {
                        if (x.ViewType == ViewType.Elevation)
                        {
                            // add view to list
                            returnList.Add(x);
                        }
                    }
                }
            }

            return returnList;
        }

        internal static List<View> GetAllViewsByCategory(Document curDoc, string catName)
        {
            List<View> m_colViews = GetAllViews(curDoc);

            List<View> m_returnList = new List<View>();

            foreach (View curView in m_colViews)
            {
                string viewCat = GetParameterValueByName(curView, "Category");

                if (viewCat == catName)
                {
                    if (curView.IsTemplate == false)
                        m_returnList.Add(curView);
                }
                    
            }

            return m_returnList;
        }

        #endregion

        #region Viewports

        internal static List<Viewport> GetAllViewports(Document curDoc)
        {
            //get all viewports
            FilteredElementCollector m_vpCollector = new FilteredElementCollector(curDoc);
            m_vpCollector.OfCategory(BuiltInCategory.OST_Viewports);

            //output viewports to list
            List<Viewport> m_vpList = new List<Viewport>();
            foreach (Viewport curVP in m_vpCollector)
            {
                //add to list
                m_vpList.Add(curVP);
            }

            return m_vpList;
        }

        #endregion

        #region View Templates

        public static List<View> GetAllViewTemplates(Document curDoc)
        {
            List<View> returnList = new List<View>();
            List<View> viewList = GetAllViews(curDoc);

            //loop through views and check if is view template
            foreach (View v in viewList)
            {
                if (v.IsTemplate == true)
                {
                    //add view template to list
                    returnList.Add(v);
                }
            }

            return returnList;
        }

        public static List<string> GetAllViewTemplateNames(Document m_doc)
        {
            //returns list of view templates
            List<string> viewTempList = new List<string>();
            List<View> viewList = new List<View>();
            viewList = GetAllViews(m_doc);

            //loop through views and check if is view template
            foreach (View v in viewList)
            {
                if (v.IsTemplate == true)
                {
                    //add view template to list
                    viewTempList.Add(v.Name);
                }
            }

            return viewTempList;
        }

        public static View GetViewTemplateByName(Document curDoc, string viewTemplateName)
        {
            List<View> viewTemplateList = GetAllViewTemplates(curDoc);

            foreach (View v in viewTemplateList)
            {
                if (v.Name == viewTemplateName)
                {
                    return v;
                }
            }

            return null;
        }

        #endregion
    }
}
