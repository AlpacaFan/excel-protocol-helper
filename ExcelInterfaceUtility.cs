using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using ExcelWindow = Microsoft.Office.Interop.Excel.Window;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using ExcelSheet = Microsoft.Office.Interop.Excel.Worksheet;
using ExcelListObject = Microsoft.Office.Interop.Excel.ListObject;
using XlCmdType = Microsoft.Office.Interop.Excel.XlCmdType;
using XlCellInsertionMode = Microsoft.Office.Interop.Excel.XlCellInsertionMode;


namespace ExcelProtocolHelper
{
    class ExcelInterfaceUtility
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);


        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
         IntPtr handle,
         uint id,
         ref Guid iid,
         out ExcelWindow excelWindow);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);


        const string MainWindowClass = "XLMAIN";
        const string DesktopWindowClass = "XLDESK";
        const string WorkbookWindowClass = "EXCEL7";
        const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        const string QueryTemplate = @"let
                        Source = OData.Feed(""::URL::"", null, [Implementation=""2.0""])
                                    in
                            Source";


        private static ISet<IntPtr> FindAllWindowsForClass(IntPtr parentHandle, string className)
        {
            HashSet<IntPtr> windows = new HashSet<IntPtr>();
            IntPtr previousWindow = IntPtr.Zero;
            IntPtr currentWindow = IntPtr.Zero;

            while (true)
            {
                currentWindow = FindWindowEx(parentHandle, previousWindow, className, null);
                if (currentWindow == IntPtr.Zero)
                    break;
                windows.Add(currentWindow);
                previousWindow = currentWindow;
            }
            return windows;
        }

        internal static ExcelWorkbook CreateNewWorkbook()
        {
            ExcelApplication application = new ExcelApplication();
            application.Visible = true;           
            return application.Workbooks.Add();
        }


        internal static void OpenLinkAsSheet(ExcelWorkbook workbook, String url)
        {
            string queryText = QueryTemplate.Replace("::URL::", url);

            //  add query            
            dynamic workbook2 =workbook ?? CreateNewWorkbook() ;
            dynamic queries = workbook2.Queries;
            dynamic query = queries.Add("ODataQuery-" + DateTime.Now.ToString("yyyyMMddhhmmss"), queryText, "");

            // add sheet with query
            ExcelSheet sheet = workbook2.Sheets.Add();

            ExcelListObject listObject = sheet.ListObjects.Add(SourceType: 0, Source: "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" + query.Name, Destination: sheet.Range["$A$1"]);

            listObject.QueryTable.CommandType = XlCmdType.xlCmdDefault;
            listObject.QueryTable.CommandText = "SELECT * FROM [" + query.Name + "]";
            listObject.QueryTable.RowNumbers = false;
            listObject.QueryTable.FillAdjacentFormulas = false;
            listObject.QueryTable.PreserveFormatting = true;
            listObject.QueryTable.RefreshOnFileOpen = false;
            listObject.QueryTable.BackgroundQuery = true;
            listObject.QueryTable.RefreshStyle = XlCellInsertionMode.xlOverwriteCells;
            listObject.QueryTable.SavePassword = false;
            listObject.QueryTable.SaveData = true;
            listObject.QueryTable.AdjustColumnWidth = true;
            listObject.QueryTable.RefreshPeriod = 0;
            listObject.QueryTable.PreserveColumnInfo = false;

            listObject.QueryTable.Refresh(true);
        }


        /// <summary>
        /// Find all open Excel workbooks in all Excel processes.
        /// </summary>
        /// <returns>List of open Excel Workbooks</returns>
        /// <remarks>
        /// This method might not be that friendly because it keeps references open. Excel might not close until finalizers are forced.
        /// </remarks>
        internal static IList<ExcelWorkbook> GetAllOpenWorkbooks()
        {
            HashSet<IntPtr> windows = new HashSet<IntPtr>();
            List<ExcelWorkbook> workbooks = new List<ExcelWorkbook>();
            Guid windowGuid = typeof(ExcelWindow).GUID;
            HashSet<uint> processes = new HashSet<uint>();

            ISet<IntPtr> mainWindows;

            mainWindows = FindAllWindowsForClass(IntPtr.Zero, MainWindowClass);

            // find all main excel windows
            foreach (IntPtr mainWindow in mainWindows)
            {
                // filter out multiple windows in the same process.
                uint processId;
                GetWindowThreadProcessId(mainWindow, out processId);
                if (processes.Contains(processId))
                    continue;
                processes.Add(processId);

                // find first desktop window
                IntPtr desktopWindow = FindWindowEx(mainWindow, IntPtr.Zero, DesktopWindowClass, null);
                if (desktopWindow != IntPtr.Zero)
                {
                    // find first workbook window
                    IntPtr workbookWindow = FindWindowEx(desktopWindow, IntPtr.Zero, WorkbookWindowClass, null);
                    if (workbookWindow != IntPtr.Zero)
                    {
                        ExcelWindow excelWindow = null;

                        if (AccessibleObjectFromWindow(workbookWindow, OBJID_NATIVEOM, ref windowGuid, out excelWindow) >= 0)
                        {
                            ExcelApplication excelApplication;
                            excelApplication = excelWindow.Application;

                            foreach (ExcelWorkbook workbook in excelApplication.Workbooks)
                            {
                                workbooks.Add(workbook);
                            }
                        }
                    }
                }
            }
            return workbooks;
        }

    }
}
