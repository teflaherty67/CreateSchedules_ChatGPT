#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.ExceptionServices;
using Forms = System.Windows;
using System.Reflection;
using System.Windows.Controls;
using System.Net;

#endregion

namespace CreateSchedules_ChatGPT
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document curDoc = uidoc.Document;

            // set some variables for paramter values
            string newFilter = "";

            if (Globals.ElevDesignation == "A")
                newFilter = "1";
            else if (Globals.ElevDesignation == "B")
                newFilter = "2";
            else if (Globals.ElevDesignation == "C")
                newFilter = "3";
            else if (Globals.ElevDesignation == "D")
                newFilter = "4";
            else if (Globals.ElevDesignation == "S")
                newFilter = "5";
            else if (Globals.ElevDesignation == "T")
                newFilter = "6";

            // open form
            frmCreateSchedules curForm = new frmCreateSchedules()
            {
                Width = 420,
                Height = 400,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            curForm.ShowDialog();

            // get form data and do something
            Globals.ElevDesignation = curForm.GetComboBoxElevationSelectedItem();
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

            // create & start the transaction
            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Create Schedules");

                // search for the sheet index
                ViewSchedule schedIndex = Utils.GetScheduleByNameContains(curDoc, "Sheet Index - Elevation " + Globals.ElevDesignation);

                // if not found, create one
                if (chbIndexResult == true && schedIndex == null)
                {
                    Utils.DuplicateAndRenameSheetIndex(curDoc, newFilter);
                }

                // search for the veneer calculations schedule
                ViewSchedule veneerIndex = Utils.GetScheduleByNameContains(curDoc, "Exterior Veneer Calculations - Elevation " + Globals.ElevDesignation);

                // if not found, create one
                if (chbVeneerResult == true && veneerIndex == null)
                {
                    Utils.DuplicateAndConfigureVeneerSchedule(curDoc);
                }

                // process floor area plans & schedule
                if (chbFloorResult == true)
                {
                    Utils.CreateAreaPlansAndSchedule(curDoc, "Floor", Globals.ElevDesignation + " Floor", "10-Floor Area", "-Schedule-", floorNum, typeFoundation);
                }

                // process frame area plans & schedule
                if(chbFrameResult == true)
                {
                    Utils.CreateAreaPlansAndSchedule(curDoc, "Frame", Globals.ElevDesignation + " Frame", "11-Frame Area", "-Frame Areas-", floorNum, typeFoundation);
                }

                // process attic area plan & schedules
                if (chbAtticResult == true)
                {
                    Utils.CreateAtticAreaPlanAndSchedule(curDoc, Globals.ElevDesignation + " Roof Ventilation", "12-Attic Area", "-Schedule-", levelWord, floorNum, typeAttic);

                    // search for equipement schedule
                    ViewSchedule equipmentSched = Utils.GetScheduleByNameContains(curDoc, "Roof Ventilation Equipment - Elevation " + Globals.ElevDesignation);

                    // if not found, create one
                    if (chbAtticResult == true && equipmentSched == null)
                    {
                        Utils.DuplicateAndConfigureEquipmentSchedule(curDoc);
                    }
                }

                t.Commit();
            }

            return Result.Succeeded;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
    }
}
