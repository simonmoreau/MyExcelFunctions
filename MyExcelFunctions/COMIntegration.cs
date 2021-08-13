using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelFunctions
{
    [ComVisible(false)]
    internal class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            // IntelliSenseServer.Install();
            // ExcelAsyncUtil.Initialize();
            ExcelIntegration.RegisterUnhandledExceptionHandler(
                ex => "!!! EXCEPTION: " + ex.ToString());
        }
        public void AutoClose()
        {
            // IntelliSenseServer.Uninstall();
        }
    }
}
