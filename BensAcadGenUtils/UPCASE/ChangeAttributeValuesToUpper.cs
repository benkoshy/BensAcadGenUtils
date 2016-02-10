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
using GetBlockInformation;

namespace BensAcadGenUtils.UPCASE
{
    class ChangeAttributeValuesToUpper
    {
        // What does this class do?
                
        ObjectIdCollection selectionSetIds;

        SelectionSet ss;

        public ChangeAttributeValuesToUpper()
        {
            // gets only block inserts
            SelectionFilter sf = GetSelectionFilter();

            // now we want a selection set out of this
            ss = BlockUtility.GetSS(sf);

            if (ss != null)
            {
                selectionSetIds = BlockUtility.GetAllIDsFromSS(ss);
            }
        }

        public ChangeAttributeValuesToUpper(SelectionSet ss)
        {
            if (ss != null)
            {
                selectionSetIds = BlockUtility.GetAllIDsFromSS(ss);
            }
        }

        private SelectionFilter GetSelectionFilter()
        {
            TypedValue[] acTypeValAr = new TypedValue[1] { 
                                new TypedValue((int)DxfCode.Start, "INSERT")                              
                                };

            // the instantiating of a selection filter which selects on certain layers and which selects only lines
            SelectionFilter sf = new SelectionFilter(acTypeValAr);

            return sf;        
        }

        public void SetAllAttributesToUpper()
        {
            // go through the entire SelectionSet Object ID collection
            foreach (ObjectId id in selectionSetIds)
            {
                BlockReference br = BlockUtility.GetBlockReference(id);

                if ( br!=null)
                {
                    int attributeCount = br.AttributeCollection.Count;

                    for (int i = 0; i < attributeCount; i++)
                    {
                        string av = BlockUtility.GetAV(id, i);
                        av = av.ToUpper();
                        BlockUtility.SetAV(id, i, av);
                    }
                }
            }
        }

        public void SetFirstAttributeToUpper()
        {
            // setting only the first attribute to upper case.
            foreach (ObjectId id in selectionSetIds)
            {
                 BlockReference br = BlockUtility.GetBlockReference(id);

                if ( br!=null)
                {
                    int attributeCount = br.AttributeCollection.Count;

                    if (attributeCount > 0)
	                {
		                string av = BlockUtility.GetAV(id, 0);
                        av = av.ToUpper();
                        BlockUtility.SetAV(id, 0, av);
	                }
                }
            }
        }


    }
}
