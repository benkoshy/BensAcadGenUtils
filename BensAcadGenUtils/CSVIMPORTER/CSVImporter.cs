using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Linq;
using Autodesk.AutoCAD.Windows;




namespace Import_CSV_to_Autocad
{
    class CSVImporter
    {
        // this program reads a CSV file and then puts it all into an autocad file
        public CSVImporter()
        {
            ImportToCad();
        }

        private void ImportToCad()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                
                if (bt != null)
                {
                    using (DocumentLock docLoc = doc.LockDocument())
                    {
                        BlockTableRecord btrModelSpace = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


                        /// Get the file etc

                        OpenFileDialog ofd = new OpenFileDialog("Select where the CSV File is located", null, "csv; txt", "CSV File - Please select where it is located", OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);
                        System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
                        if (dr == System.Windows.Forms.DialogResult.OK)
                        {
                            ed.WriteMessage("\nFile selected was \"{0}\".", ofd.Filename);



                            System.IO.StreamReader sr = new System.IO.StreamReader(ofd.Filename, true);

                            // ignore the first three rows
                            sr.ReadLine();
                            sr.ReadLine();
                            sr.ReadLine();
                            string[] headers = sr.ReadLine().Split(',');
                                                        
                            // creating a table and setting the style
                            Table tb = new Table();
                            tb.TableStyle = tb.TableStyle;

                            // inserting columns and setting row height                           
                            tb.InsertColumns(1, 10, 2);                            
                            tb.InsertRows(1, 10, 1);
                            tb.Rows[0].Height = 10;
                            for (int i = 0; i < 3; i++)
                            {
                                tb.Columns[i].Width = 45;
                            }

                            //// set up header
                            //tb.Cells[0, -1].Style = "Header";
                            //tb.Cells[1, -1].Style = "Title";

                            // add the contents of the title cell
                            Cell tc = tb.Cells[0, 0];
                            tc.Contents.Add();
                            tc.Contents[0].TextHeight = 5;
                            tc.Contents[0].TextString = "Well Locations";

                            for (int i = 0; i < 3; i++)
                            {
                                Cell c = tb.Cells[1, i];
                                c.Contents.Add();
                                c.Contents[0].TextHeight = 5;
                                c.Contents[0].TextString = headers[i];
                            }


                            //  add the rest of the contents into the table
                            while (!sr.EndOfStream)
                            {                                
                                string[] rows = sr.ReadLine().Split(',');
                                int currentRowNumber = tb.Rows.Count;
                                int rowNumberToBeAdded = currentRowNumber + 1;
                                tb.InsertRows(currentRowNumber, 10, 1);

                                for (int i = 0; i < 3; i++)
                                {
                                    Cell addingInfo = tb.Cells[currentRowNumber, i];
                                    addingInfo.Contents.Add();
                                    addingInfo.Contents[0].TextHeight = 5;

                                    if (i <2)
                                    {
                                        double wellLocation = 0;
                                        bool didParseSucceed = double.TryParse(rows[i], out wellLocation);

                                        if (!didParseSucceed)
                                        {
                                            System.Windows.Forms.MessageBox.Show("Error: could not convert one of the numbers into a double. The location data is corrupted for well " + rows[2]);
                                        }
                                        if (wellLocation == 0)
                                        {
                                            System.Windows.Forms.MessageBox.Show("Warning: Well Location is at 0 for this well. It may be because we failed to convert a string to a double propertly. Please double check that the CSV file does indeed contain a zero location at this well: " + rows[2]);
                                        }

                                        addingInfo.Contents[0].TextString = string.Format("{0:f3}", wellLocation);
                                    }
                                    else
                                    {
                                        addingInfo.Contents[0].TextString = rows[i];
                                    }
                                    
                                }
                            }

                            PromptPointResult pr = ed.GetPoint("Please click where you want the table to go.");
                            if (pr.Status == PromptStatus.OK)
                            {
                                tb.Position = pr.Value;

                                tb.GenerateLayout();

                                btrModelSpace.AppendEntity(tb);
                                tr.AddNewlyCreatedDBObject(tb, true);

                                // System.Windows.Forms.MessageBox.Show("Please check that the first and last records are here. Cross check with the CSV file.");
                            }
                        }
                    }
                    tr.Commit();                    
                }                
            }
        }

        private string GetFileName()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;


            string filePath = db.Filename;
            

            return filePath;

        }
    }
}
