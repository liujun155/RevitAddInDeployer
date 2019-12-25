using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.DB.Events;

namespace TestApp
{
    [Transaction(TransactionMode.Automatic)]
    public class CsApp : IExternalApplication
    {
        public static string AppAddInFile = typeof(CsApp).Assembly.Location;

        public static string AddInPath = Path.GetDirectoryName(AppAddInFile);

        public static string CmdAddInFile = Path.Combine(AddInPath, "TestCmd.dll");

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            RibbonPanel panel = application.CreateRibbonPanel("TestApp");

            PushButton btn = panel.AddItem(new PushButtonData("btn_App", "SayHello", CmdAddInFile, "TestCmd.CsCmd")) as PushButton;

            return Result.Succeeded;
        }

    }
}
