using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using WebSocketSharp;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //lets start this thing
            using (var ws = new WebSocket("ws://dragonsnest.far/Laputa"))
            {
                ws.OnMessage += (sender, e) =>
                    Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
                    Excel.Range firstRow = activeWorksheet.get_Range("A1");
                    firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
                    newFirstRow.Value2 = "This text was added by using code" + e.data;

                ws.Connect();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //disconnect
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
