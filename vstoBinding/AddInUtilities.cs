﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace vstoBinding
{
    [ComVisible(true)]
    public interface IAddInUtilities
    {
        void ImportData();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : IAddInUtilities
    {
        // This method tries to write a string to cell A1 in the active worksheet.
        public void ImportData()
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            if (activeWorksheet != null)
            {
                Excel.Range range1 = activeWorksheet.get_Range("A1", System.Type.Missing);
                range1.Value2 = "This is my data";
            }
        }
    }
}
