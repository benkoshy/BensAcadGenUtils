using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using GetBlockInformation;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.Colors;




namespace GetBlockInformation
{   

    /// <summary>
    /// Edit this whole class so that the blocks which START with the name are out and it is rather the blocks which match exactly which make it through
    /// </summary>
    
    class BlockUtility
    {
        // GetSSFilter – copy this code: creates a SS
        // GetSS Out: SS; In: SelectionFilter (AssumeWorks)
        // GetSS Out: SS; In: SelectionFilter, PromptSelectionOptions
        // GetAllIDsFromSS Out: all Ids of objects in SS (ObjctIDCollection); In: SS
        // GetAV Out: Attribute value;  In: BlockReferenceID, attributeName  (WORKS)
        // GetAV Out: attribute Value, In: BlockReferenceID, attributeIndex   (WORKS)
        // GetAV out: attribute value, In BlockName, attributeName
        // GetAV out: attribute value, In: blockName, attributeName, bool showError?
        // GetAT (get attribute tag): out: attributeTag (string); In: BRid, attributeIndex
        // GetBRids Out: ObjectIDCollectionReferenceIDs; In: BlockName    (ModelSpaceVersion)
        // GetBrId      Out: BlockId, In: BlockName 
        // Works: Tested!
        // GetBRids Out: ObjectIDCollectionReferenceIDs; In: BlockName, selectionset (SsetVersion)
        // Works: Tested!
        // GetAVsandBRidsFromBRids Out:  Dictionary(RefIDs, attributeValues); In (BrefIds, attributename, blockname) 
        // GetAVsandBRidsFromMS Out: Dictionary(RefIDs, attributeValues); In (attributeName, blockName) --- ModelSpaceVersion Works: Tested!
        // GetAVsandBRidsFromSS Out: Dictionary(RefIDs, attributeValues); In (selectionSet, attributename, blockname) – selectionSetVersion  Works: Tested!
        // WriteAVtoBlock Out: void; In: BlockRefID, attributeName (i.e. Tag), New attribute Value
        // SetAVs Out: void In: attributetag/attributeName, attributeValue, blockname - changes the attribute values of all blocks for a particular tag
        // SetAV Out: void; In: ObjectID, attributetag/attributeName, attributeValue
        // SetAV Out: void; In: ObjectID, attributetagIndex (int), attributeValue (string)
        // AreObjectsInPaperSpace returns bool - depending on whether there are objects in the paper space.        
        // GetInsertionPointOfBlockReference -- the blockReference numbers that go in - must actually exist. Out: Point3d, In: blockRefID, 
        // GetInsertionPoint Out: Point3d, In: blockRefID, attributeName
        // GetInsertionPoints Out: List<Point3d>, In: blockRefIDs (collection), attributeName
        
        // TESTING WHERE POINTS LOCATED
        // TestDrawLine  In: Line Out: void (draws lines in ModelSpace)
        // TestDrawCircle In: centre Out: void (draws a circle in the ModelSpace)

        // Working with TEXT and Grids (from Left to Right (or bottom to top?)): 
        //              (A) ORDEROBJECTS put in a selection set and whether the objects are ordered left to right or
        //              or bottom to top (true if horizontal) and you'll get out a List of 
        //              the x/y coordinates (depending on what the true false (horizontal/vertical)
        //              boolean is) as well as the actual length of the dimension (integer):
        //              Out: SortedList<double, int> In: SelectionSet, bool rightToLeft? true if horizontal (or bottom to top)?
        //              (B) GetTextPosition; In: SelectionSet, Out: List<Point3d> (selection set objects must be either DBtext or Mtext)
        //                  differencesInPositions: In: Dictionary<ObjectId, DBText>, bool (rightToLeft?); OUT: List<int> --> The distance between lines (which are located in the DBText value of the dictionary, are not put in an integer List).
        // GetFolderPath: Out: String(path of a folder), in: void (gets the path to a folder. if no path is successfully selected then the "MyDocuments" folder is returned"
                                                                            // Need to write the following:
                                                                            // Need to write this: Set AV given blockname and attribute tag.
        // GetPoint; In: String(instructions), Out: Point3d ***Double check to see if Point is not (0,0,0)
        // ReturnOtherEndOfLine In: Point3d, rawLine, Out: Point3d (other end of line)
        // IsPointInPolygon In: List<Point3d> polygon, Point3d point, Out: true/false

        // GetHandle In: BlockReferenceID; Out: handle        
        // GetxOrYOffset In: Point3dOrigin, Point3dOffsetPoint, bool offsettingXvalues, Out: xorYvalue offset at the case may be
        // GetxOrYOffsetPoint In: Point3dOrigin, xOryOffsetAmount, bool offsettingXvalues, Out: Point3d (OffsetPoint)
        // IsBlockReference In: ObjectID, Out: true or false bool (whether the objectID was a blockreference).
        // HighlightLine Out: Void (highlights Line) In: (Ln, bool: to highlight or unhighlight)
        // HighlightBlock Out: Void (highlights block) In: (Objectid (blockreference object id), bool: to highlight or unhighlight)
        // GetBlockReference In: BRid, Out: BlockReference, if the objectID passed in is not a block reference, an exception is thrown.
        // IsLineEqual In: Ln1, Ln2, tolerance Out: bool
        // IsBlockReferenceDuplicate In: BlockReference1, BlockReference2 Out: bool 

        // IsLineDuplicate In: ln1, ln2, Tolerance; Out: bool (tests to see whether two lines are the same with a tolerance
        // DeleteEntitesonLayer - deletes all entities on a specified layer

        public static void CheckLayerAndCreate(string layerName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                    // create a layer if none already exists
                    if ( !lt.Has(layerName))
                    {
                        LayerTableRecord ltr = new LayerTableRecord();

                        // set its properties
                        ltr.Name = layerName;

                        ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 256);

                        // add the layer to the layer table

                        lt.UpgradeOpen();

                        ObjectId ltId = lt.Add(ltr);

                        tr.AddNewlyCreatedDBObject(ltr, true);

                        tr.Commit();
                    }
                }
                
            }
        }

        public static bool AreObjectsInPaperSpace()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                    // get the first layout
                    BlockTableRecord btrLayout = tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;

                    int count = 0;
                    foreach (ObjectId id in btrLayout)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public static void  DeleteEntitiesOnLayer(string layerName, SelectionFilter sf){
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptSelectionResult psr = ed.SelectAll(sf);

            if (psr.Status == PromptStatus.OK)
            {
                ObjectIdCollection ids = new ObjectIdCollection(psr.Value.GetObjectIds());

                using (DocumentLock docLock = doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId id in ids)
                        {
                            Entity en = tr.GetObject(id, OpenMode.ForWrite) as Entity;

                            if (en!=null)
                            {
                                en.Erase();
                            }
                        }

                        tr.Commit();
                    }
                }
            }
        }

        public static bool IsLineDuplicate(Line ln1, Line ln2, Tolerance tol)
        {
            return
                (ln1.StartPoint.IsEqualTo(ln2.StartPoint, tol) || ln1.StartPoint.IsEqualTo(ln2.EndPoint, tol)) &&
                (ln1.EndPoint.IsEqualTo(ln2.EndPoint) || ln1.EndPoint.IsEqualTo(ln2.StartPoint));
        }

        public static bool IsBlockReferenceDuplicate(BlockReference blk1, BlockReference blk2, Tolerance tol)
        {            
            return
                blk1.OwnerId == blk2.OwnerId &&
                blk1.Name == blk2.Name &&
                blk1.Layer == blk2.Layer &&
                Math.Round(blk1.Rotation, 2) == Math.Round(blk2.Rotation, 2) &&
                blk1.Position.IsEqualTo(blk2.Position, tol) &&
                blk1.ScaleFactors.IsEqualTo(blk2.ScaleFactors, tol);
        }

        public static Point3d GetAttributeReferencePosition(BlockReference br, string tagName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            Point3d insertionPoint = new Point3d(0, 0, 0);

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in br.AttributeCollection)
                    {
                        DBObject tagObject = tr.GetObject(id, OpenMode.ForRead) as DBObject;

                        if (tagObject is AttributeReference)
                        {
                            AttributeReference tagReference = (AttributeReference)tagObject;

                            // check the tag's name.

                            if (tagReference.Tag.ToUpper() == tagName)
                            {
                                insertionPoint = tagReference.Position;
                            }
                        }
                    }

                    tr.Commit();
                }
            }

            // remove once you've checked it's working
            // BlockUtility.TestCircle(insertionPoint, 150);

            return insertionPoint;
        }

        public static BlockReference GetBlockReference(ObjectId BRid)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            BlockReference br;

            // check the current value of the br and see if it is null - it should not be

            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                br = tr.GetObject(BRid, OpenMode.ForRead) as BlockReference;

                if (br != null)
                {
                    return br;
                }
                // throws an exception if the point does not have any value
                else
                {
                    throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "The block reference returned is null - probably because the objectID you passed into the GetBlockReference method was not a blockReference.");
                }

            }

            return br;
        }

        public static bool IsLine(ObjectId id)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    DBObject obj = tr.GetObject(id, OpenMode.ForRead) as DBObject;

                    Line ln = obj as Line;

                    if (ln != null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                }
            }
        }

        public static bool IsBlockReference(ObjectId id)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    DBObject obj = tr.GetObject(id, OpenMode.ForRead) as DBObject;

                    BlockReference br = obj as BlockReference;

                    if (br!= null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;   
                    }
                
                } 
            }
        }

        public static Point3d GetOffsetPoint(Point3d origin, double xOffset, double yOffset)
        {
            Point3d offsetPoint = new Point3d(origin.X + xOffset, origin.Y + yOffset, 0 );
            
            return offsetPoint;
        }

        public static double GetxOrYOffset(Point3d origin, Point3d insertionPoint, bool GetXOffset)
        {
            if (GetXOffset)
            {
                double xOffSet = insertionPoint.X - origin.X;
                return xOffSet;
            }
            else
            {
                double yOffset = insertionPoint.Y - origin.Y;
                return yOffset;
            }

        }

        public static Handle GetHandle(ObjectId blockID)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            Handle h = new Handle();

            
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                DBObject block = tr.GetObject(blockID, OpenMode.ForRead) as DBObject;                           

                if (block != null)
                {
                    h = block.Handle;
                }
            }

            return h;
        }

        public static Line GetLine(ObjectId lineID)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            Line ln;

            // check the current value of the br and see if it is null - it should not be

            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                ln = lineID.GetObject(OpenMode.ForRead) as Line;

                if (ln != null)
                {
                    return ln;
                }
                // throws an exception if the point does not have any value
                else
                {
                    throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "The block reference returned is null - probably because the objectID you passed into the GetBlockReference method was not a blockReference.");
                }

            }

            return ln;
        }

        ///////////the below has proven obsolete
        //////public static BlockReference GetBlockReference(ObjectId BRid)
        //////{
        //////    Document doc = Application.DocumentManager.MdiActiveDocument;
        //////    Editor ed = doc.Editor;
        //////    Database db = doc.Database;

        //////    BlockReference br;

        //////    // check the current value of the br and see if it is null - it should not be

        //////    using (Transaction tr = doc.TransactionManager.StartTransaction())
        //////    {
        //////        br = BRid.GetObject(OpenMode.ForRead) as BlockReference;

        //////        if (br != null)
        //////        {
        //////            return br;
        //////        }
        //////        // throws an exception if the point does not have any value
        //////        else
        //////        {
        //////            throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "The block reference returned is null - probably because the objectID you passed into the GetBlockReference method was not a blockReference.");
        //////        }

        //////    }

        //////    return br;
        //////}

        public static Point3d GetInsertionPointOfBlockReference(ObjectId BRid)
        {            
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            Point3d insertionPoint = new Point3d();


            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                BlockReference br = BRid.GetObject(OpenMode.ForRead) as BlockReference;

                if (br != null)
                {
                    insertionPoint = br.Position;
                }
            }

            // throws an exception if the point does not have any value
            if (insertionPoint.IsEqualTo(new Point3d()))
            {
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.InvalidDxf3dPoint, "The block reference does not have an insertion point. Probably the objectID you passed into the GetInsertionPointOfBlockReference method is not in fact a blockReference id.");
            }

            return insertionPoint;
        }

        public static bool IsPointInPolygon(List<Point3d> polygon, Point3d point)
        {
            bool isInside = false;
            for (int i = 0, j = polygon.Count - 1; i < polygon.Count; j = i++)
            {
                if (((polygon[i].Y > point.Y) != (polygon[j].Y > point.Y)) &&
                (point.X < (polygon[j].X - polygon[i].X) * (point.Y - polygon[i].Y) / (polygon[j].Y - polygon[i].Y) + polygon[i].X))
                {
                    isInside = !isInside;
                }
            }
            return isInside;
        }

        private Point3d ReturnOtherEndOfLine(Point3d point, Line rawLine)
        {
            // Put in a line and a point which is on the line and the other end will be returned.
            if (point == rawLine.StartPoint)
            {
                return rawLine.EndPoint;
            }
            else
            {
                return rawLine.EndPoint;
            }
        }
                
        public static Point3d GetPoint(string userInstructions)
        {
            // asks the user to select a point
            // throws an exception if he cancels
            // repeats till he gets it right
            // check if point is (0,0,0)

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptPointResult pPtRes = null;
            PromptPointOptions pPtOpts = new PromptPointOptions("");

            Point3d ptStart = new Point3d(0,0,0);
                      
            // Prompt for the start point
            pPtOpts.Message = userInstructions;
            pPtRes = doc.Editor.GetPoint(pPtOpts);

            // Exit if the user presses ESC or cancels the command
            if (pPtRes.Status == PromptStatus.Cancel) throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.UserBreak, "ESCAPE pressed");

            if (pPtRes.Status == PromptStatus.OK)
            {
                ptStart = pPtRes.Value;
            }
            
            if (ptStart.X == 0 && ptStart.Y == 0 && ptStart.Z == 0)
            {
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.UserBreak, "Looks like there was an error when asked to click a point");
            }

            return ptStart;
        }
        
        public static string GetFolderPath()
        {
            //preparing the stuff to print to the text file
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.Description = "Custom Description";
            string sSelectedPath;
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                sSelectedPath = fbd.SelectedPath;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The file will be saved under 'My Documents' and will be called 'Line Differences.txt'");
                sSelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            return sSelectedPath;
        }

        public static List<int> differencesInPositions(Dictionary<ObjectId, DBText> IdsAndObjectsSorted)
        {
            
            // this takes as inputs the objectIDs and the objects themselves (i.e. the dimensions selected) and then
            // calculates the differences between the dimensions
            
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            List<int> differenceInCoordinatePositions = new List<int>();
            List<DBText> dimensions = new List<DBText>();

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in IdsAndObjectsSorted.Keys)
                    {
                        DBText text = tr.GetObject(id, OpenMode.ForRead) as DBText;
                        dimensions.Add(text);
                    }
                }
            }

            for (int i = 0; i < IdsAndObjectsSorted.Count; i++)
            {
                if (Regex.IsMatch(dimensions[i].TextString, @"^\d+$"))
                {
                    int difference = Convert.ToInt32(dimensions[i].TextString);

                    differenceInCoordinatePositions.Add(difference);       
                }         
            }
            return differenceInCoordinatePositions;
        }

        public static List<Point3d> GetTextPositions(SelectionSet ss)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            List<Point3d> textPositions = new List<Point3d>();

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject so in ss)
                    {
                        DBText dbtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as DBText;

                        MText mtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as MText;

                        // ok so let's add the db text if it's not null
                        if (dbtext != null)
                        {
                            textPositions.Add(dbtext.Position);                           
                        }

                        // but if it's an mtext then add that
                        if (mtext != null)
                        {
                            textPositions.Add(mtext.Location);                           
                        }
                    }
                }
            }

            return textPositions;
        }
        
        public static SortedList<double, int> OrderObjects(SelectionSet ss, bool rightToLeft_one_OrBottomToTop_zero) 
        {
            // gives you the DBText and ids ordered from right to left or top to bottom as the case may be
            // some dimensions might be mtext. some might be dbtext. we have to account for this.
            // we are either going in the x direction or the y direction. we have to account for this as well.
                        
            Dictionary<Point3d, int> mtextDbTextPosition_Length = new Dictionary<Point3d, int>();
            
            SortedList<double, int> xOrYPosition_DimensionLength = new SortedList<double, int>();
                       

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject so in ss)
                    {
                        DBText dbtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as DBText;

                        MText mtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as MText;

                        // ok so let's add the db text if it's not null
                        if (dbtext != null)
                        {
                            // check if it's a legitimate number, if it is, the continue on
                            if (!Regex.IsMatch(dbtext.TextString, @"^\s*\d+\s*$"))
                            {
                                // System.Windows.Forms.MessageBox.Show("Dimension " + dbtext.TextString + " is invalid. Pls check it");
                            }
                            else
                            {                               
                                mtextDbTextPosition_Length.Add(dbtext.Position, Convert.ToInt32(dbtext.TextString));
                            }
                        }

                        // but if it's an mtext then add that
                        if (mtext != null)
                        {
                            // check if it's a legitimate number, if it is, the continue on
                            if (!Regex.IsMatch(mtext.Text, @"^\s*\d+\s*$"))
                            {
                                // System.Windows.Forms.MessageBox.Show("Dimension " + dbtext.TextString + " is invalid. Pls check it");
                            }
                            else
                            {
                                mtextDbTextPosition_Length.Add(mtext.Location, Convert.ToInt32(mtext.Text));
                            }                           
                        }                        
                    }
                }                
            }                    

            // sorts from left to right, increasing the x axis
            if (rightToLeft_one_OrBottomToTop_zero == true )
            {
                if (mtextDbTextPosition_Length != null)
                {
                    foreach (KeyValuePair<Point3d, int> item in mtextDbTextPosition_Length)
                    {
                        xOrYPosition_DimensionLength.Add(item.Key.X, item.Value);
                    } 
                }
            }
            else
            {
                if (mtextDbTextPosition_Length != null)
                {
                    foreach (KeyValuePair<Point3d, int> item in mtextDbTextPosition_Length)
                    {
                        xOrYPosition_DimensionLength.Add(item.Key.Y, item.Value);
                    }
                }               
            }      
            
            return xOrYPosition_DimensionLength;
        }
        
        public static SortedList<double, string> OrderObjects_string(SelectionSet ss, bool rightToLeft_one_OrBottomToTop_zero)
        {
            // some dimensions might be mtext. some might be dbtext. we have to account for this.
            // we are either going in the x direction or the y direction. we have to account for this as well.

            Dictionary<Point3d, string> mtextDbTextPosition_Label = new Dictionary<Point3d, string>();

            SortedList<double, string> xOrYPosition_DimensionLength = new SortedList<double, string>();


            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject so in ss)
                    {
                        DBText dbtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as DBText;

                        MText mtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as MText;

                        // ok so let's add the db text if it's not null
                        if (dbtext != null)
                        {
                            if (mtextDbTextPosition_Label.ContainsKey(dbtext.Position))
                            {
                                System.Windows.Forms.MessageBox.Show("Labels must have differing positions on either the X and Y axis. Label " + dbtext.TextString + "is positioned on the same axis as another label. Please move one of the labels");
                            }
                            else
                            {
                                mtextDbTextPosition_Label.Add(dbtext.Position, dbtext.TextString);     
                            }                           
                                                   
                        }

                        // but if it's an mtext then add that
                        if (mtext != null)
                        {
                            if (mtextDbTextPosition_Label.ContainsKey(mtext.Location))
                            {
                                System.Windows.Forms.MessageBox.Show("two labels cannot have the same position. Label " + mtext.Text + "is positioned exactly the same as another label. Please move one of the labels");
                            }
                            else
                            {
                                mtextDbTextPosition_Label.Add(mtext.Location, mtext.Text);                            
                            }                            
                        }
                    }
                }
            }

            // sorts from left to right, increasing the x axis
            if (rightToLeft_one_OrBottomToTop_zero == true)
            {
                if (mtextDbTextPosition_Label != null)
                {
                    foreach (KeyValuePair<Point3d, string> item in mtextDbTextPosition_Label)
                    {
                        if (!xOrYPosition_DimensionLength.ContainsKey(item.Key.X))
                        {
                            xOrYPosition_DimensionLength.Add(item.Key.X, item.Value);
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("Same X coordinate for Label " + item.Value + "./n/n It is positioned exactly the same as another label. Please move one of the labels.");
                        }
                    }
                }
            }
            else
            {
                if (mtextDbTextPosition_Label != null)
                {
                    foreach (KeyValuePair<Point3d, string> item in mtextDbTextPosition_Label)
                    {
                        // only add if it doesn't contain a key
                        if (!xOrYPosition_DimensionLength.ContainsKey(item.Key.Y))
                        {
                            xOrYPosition_DimensionLength.Add(item.Key.Y, item.Value);
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("Same Y coordinate for Label " + item.Value + "\n\n It is positioned exactly the same as another label. Please move one of the labels.");
                        }
                        
                    }
                }
            }

            return xOrYPosition_DimensionLength;
        }

        //public static Dictionary<ObjectId, DBText> OrderObjects(SelectionSet ss, bool rightToLeft_one_OrBottomToTop_zero)
        //{
        //    // create an unordered list
        //    List<DBText> UnOrderedListBRs = new List<DBText>();
        //    List<DBText> SortedBRs;
        //    Dictionary<ObjectId, DBText> IdsAndObjectsSorted = new Dictionary<ObjectId, DBText>();

        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Editor ed = doc.Editor;
        //    Database db = doc.Database;

        //    using (DocumentLock docLock = doc.LockDocument())
        //    {
        //        using (Transaction tr = db.TransactionManager.StartTransaction())
        //        {
        //            foreach (SelectedObject so in ss)
        //            {
        //                DBText dbtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as DBText;

        //                MText mtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as MText;

        //                if (dbtext != null)
        //                {
        //                    UnOrderedListBRs.Add(dbtext);
        //                }


        //                if (mtext != null)
        //                {
        //                    UnOrderedListBRs.Add(dbtext);
        //                }


        //            }
        //        }
        //    }

        //    // sorts from left to right, increasing the x axis
        //    if (rightToLeft_one_OrBottomToTop_zero == true)
        //    {
        //        SortedBRs = UnOrderedListBRs.OrderBy(o => o.Position.X).ToList();
        //    }
        //    else
        //    {
        //        SortedBRs = UnOrderedListBRs.OrderBy(o => o.Position.Y).ToList();
        //    }


        //    using (DocumentLock docLock = doc.LockDocument())
        //    {
        //        using (Transaction tr = db.TransactionManager.StartTransaction())
        //        {
        //            foreach (DBText text in SortedBRs)
        //            {
        //                IdsAndObjectsSorted.Add(text.Id, (tr.GetObject(text.Id, OpenMode.ForRead) as DBText));
        //            }
        //        }
        //    }

        //    return IdsAndObjectsSorted;
        //}

        public static void HighlightLine(Line ln, bool highlight)
        {
            // Get the current document and database
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Start a transaction
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Line lineToHighlight = tr.GetObject(ln.Id, OpenMode.ForWrite) as Line;

                    if (highlight)
                    {
                        lineToHighlight.Highlight();
                    }
                    else
                    {
                        lineToHighlight.Unhighlight();
                    }

                    // Save the new object to the database
                    tr.Commit();
                }
            }
        }

        public static void HighlightBlock(ObjectId id, bool highlight)
        {
            // Get the current document and database
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Start a transaction
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity en = tr.GetObject(id, OpenMode.ForWrite) as Entity;

                    if (highlight)
                    {
                        en.Highlight();
                    }
                    else
                    {
                        en.Unhighlight();
                    }

                    // Save the new object to the database
                    tr.Commit();
                }
            }
        }

        public static void TestCircle(Point3d circleCentre, double radius)
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                    OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                    OpenMode.ForWrite) as BlockTableRecord;

                    Circle c1 = new Circle();
                    c1.Center = new Point3d(circleCentre.X, circleCentre.Y, 0);
                    c1.Radius = radius;

                    // Add the new object to the block table record and the transaction
                    acBlkTblRec.AppendEntity(c1);
                    acTrans.AddNewlyCreatedDBObject(c1, true);


                    // Save the new object to the database
                    acTrans.Commit();
                }
            }
        }

        public static void TestLines(Line ln)
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                    OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                    OpenMode.ForWrite) as BlockTableRecord;

                    // Add the new object to the block table record and the transaction
                    acBlkTblRec.AppendEntity(ln);
                    acTrans.AddNewlyCreatedDBObject(ln, true);
                    
                    // Save the new object to the database
                    acTrans.Commit();
                }
            }
        }

        public static string GetAV(string blockName, string attributeName)
        {
            string attributeValue = "";

            ObjectId blockBRid = GetBRid(blockName);

            if (blockBRid.IsNull)
            {
                // System.Windows.Forms.MessageBox.Show("We are searching for a block and we want the first block reference number of this block in the modelspace. We could not find a block with the block name programmed");
            }
            else
            {
                attributeValue = GetAV(blockBRid, attributeName);
            }                       

            return attributeValue;
        }

        public static string GetAV(string blockName, string attributeName, bool showError)
        {
            string attributeValue = "";

            ObjectId blockBRid = GetBRid(blockName);

            if (blockBRid.IsNull)
            {
                if (showError)
                {
                    System.Windows.Forms.MessageBox.Show("We are searching for a block and we want the first block reference number of this block in the modelspace. We could not find a block with the block name programmed"); 
                }
            }
            else
            {
                attributeValue = GetAV(blockBRid, attributeName);
            }

            return attributeValue;
        }
        
        public static void SetAVs(string blockname, string attributeTag, string attributeValue)
        {

            ObjectIdCollection idsOfBlocksWithBlockName = GetBRids(blockname);

            foreach (ObjectId id in idsOfBlocksWithBlockName)
            {
                SetAV(id, attributeTag, attributeValue.ToString());
            }  
        }

          public static void SetAV(ObjectId id, string attributeTag, string attributeValue)
          {
              try
              {
                  Document doc = Application.DocumentManager.MdiActiveDocument;
                  Editor ed = doc.Editor;
                 
                  Database db = doc.Database;

                  using (Transaction tr = doc.TransactionManager.StartTransaction())
                  {

                      using (DocumentLock doclock = doc.LockDocument() )
                      {
                          BlockReference br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;

                          if (br != null)
                          {
                              AttributeCollection arColl = br.AttributeCollection;

                              if (arColl != null)
                              {

                                  foreach (ObjectId arID in arColl)
                                  {

                                      AttributeReference ar = tr.GetObject(arID, OpenMode.ForRead) as AttributeReference;

                                      if (ar.Tag == attributeTag)
                                      {
                                          br.UpgradeOpen();
                                          ar.UpgradeOpen();
                                          ar.TextString = attributeValue;
                                      }
                                  }
                              }
                          } 
                      }
                      
                      tr.Commit();
                  }
              }
              catch (System.Exception ex)
              {
                   System.Windows.Forms.MessageBox.Show(ex.Message);                  
              }
          }

          public static void SetAV(ObjectId id, int TagIndex, string attributeValue)
          {
              try
              {
                  Document doc = Application.DocumentManager.MdiActiveDocument;
                  Editor ed = doc.Editor;

                  Database db = doc.Database;

                  using (Transaction tr = doc.TransactionManager.StartTransaction())
                  {

                      using (DocumentLock doclock = doc.LockDocument())
                      {
                          BlockReference br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;

                          if (br != null)
                          {
                              AttributeCollection arColl = br.AttributeCollection;

                              if (arColl != null)
                              {
                                  if (TagIndex < arColl.Count)
                                  {
                                      ObjectId attributeReferenceID = arColl[TagIndex];

                                      if (!attributeReferenceID.IsNull )
                                      {
                                          AttributeReference ar = tr.GetObject(attributeReferenceID, OpenMode.ForRead) as AttributeReference;

                                          br.UpgradeOpen();
                                          ar.UpgradeOpen();
                                          ar.TextString = attributeValue;                                          
                                      }
                                  }
                              }
                          }
                      }

                      tr.Commit();
                  }
              }
              catch (System.Exception ex)
              {
                  System.Windows.Forms.MessageBox.Show(ex.Message);
              }
          }              



        public static void WriteToEditor(string str)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage(str);            
        }
        
        public static void WriteAVtoBlock(ObjectId brID, string attbName, string newAttbValue)  // Successfully works - just ensure that the tag is correct - case sensitive
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    
                        BlockReference br = tr.GetObject(brID, OpenMode.ForRead) as BlockReference;

                        if (br != null)
                        {
                            AttributeCollection arColl = br.AttributeCollection;

                            if (arColl != null)
                            {

                                foreach (ObjectId arID in arColl)
                                {
                                    AttributeReference ar = tr.GetObject(arID, OpenMode.ForRead) as AttributeReference;

                                    if (ar != null)
                                    {
                                        if (ar.Tag == attbName)
                                        {
                                            ar = tr.GetObject(arID, OpenMode.ForWrite) as AttributeReference;
                                            ar.TextString = newAttbValue;
                                            ed.WriteMessage("Successfully written!");
                                        }
                                    }
                                }
                            }

                        }
                    

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());                
            }
        }

        public static void WriteAVtoBlock(ObjectId brID, int attributeIndex, string newAttbValue)  // Successfully works - just ensure that the tag is correct - case sensitive
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                using (DocumentLock docLock = doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {

                        BlockReference br = tr.GetObject(brID, OpenMode.ForRead) as BlockReference;

                        if (br != null)
                        {
                            AttributeCollection arColl = br.AttributeCollection;

                            if (arColl != null)
                            {

                                if (attributeIndex < arColl.Count)
                                {
                                    ObjectId arID = arColl[attributeIndex];

                                    AttributeReference ar = tr.GetObject(arID, OpenMode.ForRead) as AttributeReference;

                                    if (ar != null)
                                    {

                                        ar = tr.GetObject(arID, OpenMode.ForWrite) as AttributeReference;
                                        ar.TextString = newAttbValue;
                                        // ed.WriteMessage("Successfully written!");
                                    }
                                }
                                else
                                {
                                    ed.WriteMessage("Error - the index you provided for the attribute collection is out of range");
                                    System.Windows.Forms.MessageBox.Show("Error - the index you provided for the attribute collection is out of range");
                                }

                            }

                        }


                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private static SelectionFilter GetSSFilter()   
        {
            // selection filter. We filter by the name of the panel and also we want only lines
            TypedValue[] acTypeValAr = new TypedValue[4] {new TypedValue((int)DxfCode.Operator, "<or"), 
                                new TypedValue((int)DxfCode.LayerName, "PANELS"),
                                new TypedValue((int)DxfCode.LayerName, "0"),
                                new TypedValue((int)DxfCode.Operator, "or>")                                
                                };

            // the instantiating of a selection filter which selects on certain layers and which selects only lines
            SelectionFilter sf = new SelectionFilter(acTypeValAr);

            return sf;
        }
        
        public static Dictionary<ObjectId, string> GetAVsandBRidsfromSS(string blockName, string attributeName, SelectionSet ss)
        {
            ObjectIdCollection BrefIds = GetBRids(ss, blockName);
            return GetAVsandBRidsFromBRids(blockName, attributeName, BrefIds);
        }
        
        public static ObjectIdCollection GetAllIDsFromSS(SelectionSet ss) // Returns all selected object IDS in a selection set - all IDs no matter the name of the blocks etc.
        {
            ObjectIdCollection idsOfObjects = new ObjectIdCollection();

            foreach (SelectedObject so in ss)
            {
                idsOfObjects.Add(so.ObjectId);
            }

            return idsOfObjects;
        }

        public static ObjectIdCollection GetBRids(SelectionSet ss, string blockname)   // Successfully works
        {
            ObjectIdCollection idsOfSS_containsAllObjectIds = GetAllIDsFromSS(ss);              // id collection contains all IDS
            ObjectIdCollection idsInSS_containsOnlyBlockNames = new ObjectIdCollection();       // id collection which needs irrelevant components to be removed from it

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            foreach (ObjectId id in idsOfSS_containsAllObjectIds)                               // populating ID collection. We must now remove from this collection all blockRefIds which do not have blockname in them.
            {
                idsInSS_containsOnlyBlockNames.Add(id);
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in idsInSS_containsOnlyBlockNames)
                {
                    BlockReference br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    
                    if (br != null)
                    {
                        if (! (br.Name == blockname))
                        {
                            idsInSS_containsOnlyBlockNames.Remove(id);
                        }                        
                    }
                } 
            }
            
            return idsInSS_containsOnlyBlockNames;
        }

        public static SelectionSet GetSS(SelectionFilter sf, PromptSelectionOptions pso) 
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            //clear the current selection
            ObjectId[] idarrayEmpty = new ObjectId[0];
            ed.SetImpliedSelection(idarrayEmpty);
            
            PromptSelectionResult psr = ed.GetSelection(pso, sf);

            return psr.Value;
        }

        public static SelectionSet GetSS(SelectionFilter sf) // Works!! MUST CHECK if something is selected - if not, then return null please for goddsakes!!
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            //clear the current selection
            ObjectId[] idarrayEmpty = new ObjectId[0];
            ed.SetImpliedSelection(idarrayEmpty);

            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nSelect block(s)";

            PromptSelectionResult psr = ed.GetSelection(pso, sf);

            return psr.Value;            
        }

        public static Dictionary<ObjectId, string> GetAVsandBRidsFromMS(string blockname, string attributeName)
        {
            ObjectIdCollection blockRefIDs = GetBRids(blockname);
            return GetAVsandBRidsFromBRids(blockname, attributeName, blockRefIDs);
        }

        public static Dictionary<ObjectId, string> GetAVsandBRidsFromBRids(string blockname, string attributeName, ObjectIdCollection blockRefIDs)
        {
            // blockRefIDs = GetBlockReferenceIdsGivenABlockName(blockname);
            Dictionary<ObjectId, string> RefIDs_AttributeValues = new Dictionary<ObjectId, string>();
            string attributeValue = "";

            foreach (ObjectId id in blockRefIDs)
            {
                attributeValue = GetAV(id, attributeName);
                RefIDs_AttributeValues.Add(id, attributeValue);
            }

            return RefIDs_AttributeValues;
        }

        public static ObjectIdCollection GetBRids(string blockname)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ObjectIdCollection blockReferenceIDs = new ObjectIdCollection();            

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                foreach (ObjectId btrID in bt)
                {
                    BlockTableRecord btr = tr.GetObject(btrID, OpenMode.ForRead) as BlockTableRecord;

                    if (btr != null)
                    {
                        if ( btr.Name == blockname)
                        {
                            blockReferenceIDs = btr.GetBlockReferenceIds(true, true);
                        }
                    }
                }
            }

            return blockReferenceIDs;
        }

        public static ObjectId GetBRid(string blockname)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ObjectIdCollection blockReferenceIDs = new ObjectIdCollection();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                foreach (ObjectId btrID in bt)
                {
                    BlockTableRecord btr = tr.GetObject(btrID, OpenMode.ForRead) as BlockTableRecord;

                    if (btr != null)
                    {
                        if (btr.Name == blockname)
                        {
                            blockReferenceIDs = btr.GetBlockReferenceIds(true, true);
                        }
                    }
                }
            }

            if (blockReferenceIDs.Count > 0)
            {
                return blockReferenceIDs[0];    
            }
            else
            {
                return new ObjectId();
            }
        }

        public static string GetAV(ObjectId id, string attributeName) //pass in the blockreferenceID of the block
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string attributevalue = "";
             
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                    BlockReference br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;

                     if (br!= null)
                     {
                         AttributeCollection arColl = br.AttributeCollection;

                         if (arColl != null)
                         {

                             foreach (ObjectId arID in arColl)
                             {

                                 AttributeReference ar = tr.GetObject(arID, OpenMode.ForRead) as AttributeReference;

                                 if (ar.Tag == attributeName)
                                 {
                                     
                                     attributevalue = ar.TextString;
                                 }
                             }
                         } 
                     }
                 
             }
            return attributevalue;
        }

        public static string GetAV(ObjectId id, int attributeIndex) //pass in the blockreferenceID of the block
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string attributevalue = "";

            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                using (DocumentLock lock1 = doc.LockDocument())
                {
                    BlockReference br = tr.GetObject(id, OpenMode.ForWrite) as BlockReference;                   

                    AttributeCollection arColl = br.AttributeCollection;

                    if (arColl!= null)
                    {
                        ObjectId arId = arColl[attributeIndex];

                        AttributeReference ar = tr.GetObject(arId, OpenMode.ForRead) as AttributeReference;

                        if (ar != null)
                        {
                            attributevalue = ar.TextString;
                        } 
                    }
                }
            }
            return attributevalue;
        }

        public static string GetAT(ObjectId id, int attributeIndex) //pass in the blockreferenceID of the block
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string attributeTag = "";

            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                using (DocumentLock lock1 = doc.LockDocument())
                {
                    BlockReference br = tr.GetObject(id, OpenMode.ForWrite) as BlockReference;

                    AttributeCollection arColl = br.AttributeCollection;

                    if (arColl != null)
                    {
                        ObjectId arId = arColl[attributeIndex];

                        AttributeReference ar = tr.GetObject(arId, OpenMode.ForRead) as AttributeReference;

                        if (ar != null)
                        {
                            attributeTag = ar.Tag;
                        }
                    }
                }
            }
            return attributeTag;
        }


       public List<Point3d> GetInsertionPoints(ObjectIdCollection BRids, string attributeName)
       {
           List<Point3d> insertionPoints = new List<Point3d>();
           
           foreach (ObjectId BRid in BRids)
	        {               
               Point3d insertionPoint = GetInsertionPoint(BRid, attributeName);
               insertionPoints.Add(insertionPoint);
	        }

           return insertionPoints;
       }
        
        public Point3d GetInsertionPoint(ObjectId BRid, string attributeName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            Point3d insertionPoint = new Point3d();
            
            
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {               
                    BlockReference br = BRid.GetObject(OpenMode.ForRead) as BlockReference;

                    if (br != null)
                    {
                        AttributeCollection arColl = br.AttributeCollection;

                        foreach (ObjectId arID in arColl)
                        {
                            AttributeReference ar = tr.GetObject(arID, OpenMode.ForRead) as AttributeReference;                           

                            if (ar != null)
                            {
                                if (ar.Tag == attributeName)
                                {
                                    insertionPoint = ar.Position;
                                }

                            }
                        }
                    }
                }
            return insertionPoint;
            }                

        internal static string GetAVfromPaperSpaceOnly(string layoutDrawingTitleBar, string attbNameForlayoutDrawingTitleBar, bool showError)
        {
            // NOTE: that the above layoutDrawingTitleBar could also be a shopDrawingTitleBar

            string attributeValue = "";

            ObjectId blockBRid = GetBRidFromPaperSpace(layoutDrawingTitleBar);

            if (blockBRid.IsNull)
            {
                if (showError)
                {
                    System.Windows.Forms.MessageBox.Show("We are searching for a block and we want the first block reference number of this block in the modelspace. We could not find a block with the block name programmed");
                }
            }
            else
            {
                attributeValue = GetAV(blockBRid, attbNameForlayoutDrawingTitleBar);
            }

            return attributeValue;
        }

        private static ObjectId GetBRidFromPaperSpace(string layoutDrawingTitleBar) // picks up the first one
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ObjectIdCollection finalID = new ObjectIdCollection();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                foreach (ObjectId btrID in bt)
                {
                    BlockTableRecord btr = tr.GetObject(btrID, OpenMode.ForRead) as BlockTableRecord;

                    if (btr != null)        // check of the block table record is not null
                    {
                        if (btr.IsLayout)   // ensure that we are looking at things in the layout
                        {
                            foreach (ObjectId layoutID in btr)
                            {
                                // we only want to pick up blocks
                                BlockReference br = tr.GetObject(layoutID, OpenMode.ForRead) as BlockReference;

                                if (br != null)
                                {
                                    // and we want only the first one we stumble upon

                                    if (br.Name == layoutDrawingTitleBar)
                                    {
                                        return layoutID;
                                    }

                                }
                            }
                        }
                    }
                }
            }           
            return new ObjectId();            
        } // return get BRid from paperspace.


    }

    }


    

    

