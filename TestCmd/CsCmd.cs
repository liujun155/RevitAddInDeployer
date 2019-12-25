using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Utility;

namespace TestCmd
{
    [Transaction(TransactionMode.Automatic)]
    public class CsCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Tip","Hello Revit");

            return Result.Succeeded;
        }

    }
}
