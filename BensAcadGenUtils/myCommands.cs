// (C) Copyright 2016 by  
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using GetBlockInformation;
using Import_CSV_to_Autocad;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(BensAcadGenUtils.MyCommands))]

namespace BensAcadGenUtils
{

    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class MyCommands
    {
        // The CommandMethod attribute can be applied to any public  member 
        // function of any public class.
        // The function should take no arguments and return nothing.
        // If the method is an intance member then the enclosing class is 
        // intantiated for each document. If the member is a static member then
        // the enclosing class is NOT intantiated.
        //
        // NOTE: CommandMethod has overloads where you can provide helpid and
        // context menu.

        // Modal Command with localized name
        [CommandMethod("MyGroup", "MyCommand", "MyCommandLocal", CommandFlags.Modal)]
        public void MyCommand() // This method can have any name
        {
            // Put your command code here
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed;
            if (doc != null)
            {
                ed = doc.Editor;
                ed.WriteMessage("Hello, this is your first command.");

            }
        }

        // Modal Command with pickfirst selection
        [CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void MyPickFirst() // This method can have any name
        {
            PromptSelectionResult result = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection();
            if (result.Status == PromptStatus.OK)
            {
                // There are selected entities
                // Put your command using pickfirst set code here
            }
            else
            {
                // There are no selected entities
                // Put your command code here
            }
        }

        // Application Session Command with localized name
        [CommandMethod("UPCASEALL", "UPCASEALL", "UPCASEALL", CommandFlags.Modal | CommandFlags.Session)]
        public void UPCASEALL() // This method can have any name
        {
            BensAcadGenUtils.UPCASE.ChangeAttributeValuesToUpper a = new UPCASE.ChangeAttributeValuesToUpper();
            a.SetAllAttributesToUpper();
        }

        [CommandMethod("UPCASEFIRST", "UPCASEFIRST", "UPCASEFIRST", CommandFlags.Modal | CommandFlags.Session)]
        public void UPCASEFIRST() // This method can have any name
        {
            BensAcadGenUtils.UPCASE.ChangeAttributeValuesToUpper a = new UPCASE.ChangeAttributeValuesToUpper();
            a.SetFirstAttributeToUpper();
        }

        [CommandMethod("CSVImport", CommandFlags.Modal | CommandFlags.Session)]
        public void CSVImport() // This method can have any name
        {
            CSVImporter a1 = new CSVImporter();
        }

        #region Chk Double Count (Wells) program

        [CommandMethod("ChkDoubleCount", CommandFlags.UsePickSet | CommandFlags.Modal | CommandFlags.Session)]
        public void ChkDoubleCount() // This method can have any name
        {
            ChkDoubleCountAllWells();
        }

        // This program looks at all the wells in the drawing.
        // First the drawing wells must be selected prior to running - only drawing wells should be selected.
        // Then a small regex is applied so that wells with F07Rochester are renamed F7Rochester 
        // the numeral 0 is taken out immediately after the starting letter.        
        // A warning is then displayed for any wells which have duplicate names
        // A well will have a duplicate name if the ending number is the same.
        // e.g. F050 and C00050 will strike a warning saying that 50 is a duplicate

        private void ChkDoubleCountAllWells()
        {

            try
            {
                // warns the user that only wells must be selected
                // that pick first must operate otherwise it will crash.
                PrintWarning();

                // obtains selection set of wells. only wells must be selected in a pick first selection
                SelectionSet ss = GetSelectionSet();

                if (ss != null)
                {
                    // gets IDS of the wells
                    ObjectIdCollection BRids = BlockUtility.GetAllIDsFromSS(ss);


                    // preparing a list of attribute values - we will use this to check for double values
                    List<string> attributeValues = new List<string>();

                    // loops through wells and applies a regex
                    foreach (ObjectId id in BRids)
                    {
                        // Get AV - it gets the first one in the list
                        string attributeValue = BlockUtility.GetAV(id, 0);
                        string attributeValueWellNumberOnly = "";

                        // the following works on duplicate well numbers only
                        // we want the well names to be: just numbers - the names must be stripped of
                        // the first letter and any subsequent zeros will be removed
                        if (CheckIfAVNeedsCleaning(attributeValue, @"^[a-zA-Z]*0*"))
                        {
                            // gives the new attribute value
                            attributeValueWellNumberOnly = CleanAttributeValue(attributeValue, @"^[a-zA-Z]*0*", "");
                        }

                        // add attribute values to list
                        attributeValues.Add(attributeValueWellNumberOnly);


                        // the following works on just updating the attribute values of the well 
                        // we want the block to remove any zeroes after the initial few letters
                        if (CheckIfAVNeedsCleaning(attributeValue, @"^[a-zA-Z]*0+"))
                        {
                            attributeValue = CleanAttributeValue(attributeValue, @"^([a-zA-Z]*)(0+)", "$1");
                            BlockUtility.WriteAVtoBlock(id, 0, attributeValue);
                        }
                    }

                    //Prints duplicate well names to the console and via msg box
                    PrintDuplicateWellNames(attributeValues);
                }
            }
            catch (System.Exception)
            {

            }
        }

        private void PrintDuplicateWellNames(List<string> attributeValues)
        {
            List<string> uniqueWellNames = new List<string>();
            List<string> duplicateWellNames = new List<string>();

            foreach (string wellName in attributeValues)
            {
                if (!uniqueWellNames.Contains(wellName))
                {
                    uniqueWellNames.Add(wellName);
                }
                else
                {
                    duplicateWellNames.Add(wellName);
                }
            }

            // concatenate all the duplicate well names into a single string
            string concatenatedWellNames = String.Join(" ", duplicateWellNames.ToArray());

            //print the results in a message box
            if (duplicateWellNames.Count > 0)
            {
                System.Windows.Forms.MessageBox.Show("The following have more than one wells of the same name: " + concatenatedWellNames);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No duplicate well names found!");
            }



        }

        static bool CheckIfAVNeedsCleaning(string attributeValue, string pattern)
        {

            // instantiate regex object
            Regex r = new Regex(pattern, RegexOptions.IgnoreCase);

            // match regular expression against a text sting
            Match m = r.Match(attributeValue);

            if (m.Success)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static string CleanAttributeValue(string strIn, string regexPattern, string replacementString)
        {
            // Replace invalid characters with empty strings.
            try
            {
                return Regex.Replace(strIn, regexPattern, replacementString, RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1.5));
            }
            // If we timeout when replacing invalid characters, 
            // we should return Empty.
            catch (RegexMatchTimeoutException)
            {
                return String.Empty;
            }
        }

        private void PrintWarning()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            ed.WriteMessage("1. Select only wells - nothing else, and 2. select the wells before running the command, otherwise errors may result. /n The well name must be the very first attribute in the list or it won't work!");
        }

        private SelectionSet GetSelectionSet()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptSelectionResult psr = ed.SelectImplied();

            SelectionSet ss;

            if (psr.Status == PromptStatus.OK)
            {
                ss = psr.Value;
                return ss;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select some wells before running the command.");
                return ss = null;
            }


        }

        private SelectionFilter GetSelectionFilter()
        {
            // selection filter. We filter by the name of the panel and also we want only lines
            TypedValue[] acTypeValAr = new TypedValue[1] { new TypedValue((int)DxfCode.Operator, "INSERT") };

            // the instantiating of a selection filter which selects on certain layers and which selects only lines
            SelectionFilter sf = new SelectionFilter(acTypeValAr);

            return sf;
        } 
        #endregion

    }

}
