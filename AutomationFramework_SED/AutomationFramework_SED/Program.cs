using AutomationFramework_SED.BaseFiles;
using AutomationFramework_SED.Reports;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_SED
{
    class Program
    {
        static void Main(string[] args)
        {

            TestclsData objData = new TestclsData();
            objData.LoadFile(@"C:\Users\omar.feria\Downloads\P&G.xlsx", "SetupInfo");
            for (int i = 1; i <= objData.RowCount; i++)
            {
                objData.CurrentRow = i;
                //Console.WriteLine(objData.GetValueNew(2, ""));
            }

            Console.WriteLine();



            //TestclsData.FilePath = @"C:\Users\omar.feria\Downloads\P&G.xlsx";
            //TestclsData.LoadFile(@"C:\Users\omar.feria\Downloads\P&G.xlsx");
            //TestclsData.LoadSheet("SetupInfo");
           






            SLDocument sl = new SLDocument(@"C:\Users\omar.feria\Downloads\P&G.xlsx");
            sl.SelectWorksheet("Setup");

            Dictionary<string, int> _dicHeaders = new Dictionary<string, int>();
            for (int i = 1; i <= sl.GetWorksheetStatistics().NumberOfColumns; i++)
            {
                _dicHeaders.Add(sl.GetCellValueAsString(1, i), i);
            }
            Console.WriteLine(_dicHeaders["#"]);
            Console.WriteLine(_dicHeaders["Setup Item"]);
            Console.WriteLine(_dicHeaders["Response"]);

            string x;
            x = ""; 






            int RowStartIndex = sl.GetWorksheetStatistics().NumberOfRows;
            int RowStartIndex2 = sl.GetWorksheetStatistics().NumberOfColumns;
            int ColumnStartIndex = sl.GetWorksheetStatistics().StartColumnIndex;
            for (int i = 0; i <= RowStartIndex; i++)
            {
                Console.WriteLine(sl.GetCellValueAsString(i, 2));
            }


            Datasheet.FileOpen(@"C:\Users\omar.feria\Downloads\P&G.xlsx", "Setup");
            int Rows = Datasheet.RowCount;
            for (int intRow = 2; intRow <= Datasheet.RowCount; intRow++)
            {
                Datasheet.CurrentRow = intRow;
                Console.WriteLine(Datasheet.GetValue("SetupItem", ""));


            }
            
            string varUser = Datasheet.GetValue("Setup Item", "");
            string varUser2 = Datasheet.GetValue("Setup Item", "");

            Console.ReadLine();
            
        }
    }
}
