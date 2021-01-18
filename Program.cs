using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Regex = System.Text.RegularExpressions.Regex;
using System.IO;
using System.IO.Compression;
using Microsoft.Office.Interop.Excel;

namespace attempt2
{
    class Program
    {
        static void copypaste(Excel.Application excelapp, Excel.Workbook source, Excel.Workbook destination, string worksheetname, string filter, int filterstartcell, int filtercolumn, string lastcolumn, bool straightcopy)
        {
            // function takes in two excel workbooks, worksheet names and filter parameters for dynamic filtering, copying and pasting of excel sheets
            Excel.Worksheet sourceworksheet = source.Worksheets[worksheetname];

            //copies original worksheet to new workbook
            sourceworksheet.Copy(destination.Worksheets[1]);

            Excel.Worksheet destinationworksheet = destination.Worksheets[worksheetname];

            //assign max range of excel worksheet to variable
            long rows = sourceworksheet.UsedRange.Rows.Count;

            //clears contents of copied sheet but not formatting
            destinationworksheet.Range["A" + (filterstartcell + 1).ToString(), lastcolumn + (rows + 1)].Clear();

            //establishes the range of cells to be filtered and copied 
            string filterrange = "A" + filterstartcell.ToString() + ":" + lastcolumn + rows;
            string copyrange = "A" + (filterstartcell+1).ToString() +":" + lastcolumn + rows;

            try
            {
                //filters worksheet
                sourceworksheet.Range[filterrange].AutoFilter(filtercolumn, "=*" + filter + "*", Excel.XlAutoFilterOperator.xlAnd);

                //copies worksheet according one of to two different methods, according to Main method specification
                if (straightcopy == true)
                { 
                    sourceworksheet.Range[copyrange].Copy(destinationworksheet.Range[copyrange]); 
                }
                if (straightcopy == false)
                {
                    sourceworksheet.Range[copyrange].SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy();
                    destinationworksheet.Range[copyrange].PasteSpecial();
                }
                
            }

            catch (Exception)
            {
                destinationworksheet.Delete();
                Console.WriteLine("Cannot filter " + worksheetname + ", sheet removed for "+ destination.Path);
            }
        }

        static void Main(string[] args)
        {
            //static variables
            List<string> tmTerritory = new List<string>() { "SalesTerritoryID-SalesTerritoryName-SalesRepresentative", "SalesTerritoryID-SalesTerritoryName-SalesRepresentative" };
            Dictionary<string, string> rsdRegion = new Dictionary<string, string>() { { "SalesRegionName", "SalesRegionDirector" }, { "SalesRegionName", "SalesRegionDirector" } };

            //worksheets to be copied without filtering
            List<string> staticRSD = new List<string>() { "Sheet2", "Sheet3"};
            List<string> staticTM = new List<string>() { "Sheet2" , "Sheet3" , "Sheet4" };
            List<string> staticPCPTM = new List<string>() { "Sheet5", "Shhet6"};

            //assigns current date as string to cleanDate
            string rawDate = DateTime.Today.ToShortDateString();
            string cleanDate = Regex.Replace(rawDate, "[^A-Za-z0-9 ]", "");

            string Q121IC = @"C:\Users\User\Documents\workbook1.xlsx";

            //loops through each sales representative in the tmTerritory list
            foreach (var territory in tmTerritory)
            {
                Excel.Application excelapp = new Excel.Application();
                
                //hides excel operations being performed
                excelapp.Visible = false;
                excelapp.DisplayAlerts = false;

                Excel.Workbook wb1 = excelapp.Workbooks.Open(Q121IC);

                //adds new workbook
                Excel.Workbook output = excelapp.Workbooks.Add();

                //assigns first character of territory number to terr0
                char terr0 = territory[0];

                //copies different worksheets based off terr0, which denotes if which team the sales representative is part of
                if (terr0 == '2')
                { 
                    foreach (string ws in staticPCPTM)
                    {
                        wb1.Worksheets[ws].Copy(output.Worksheets[1]);
                    }
                }

                if (terr0 == '1')
                {
                    wb1.Worksheets["Sheet7"].Copy(output.Worksheets[1]);
                }

                copypaste(excelapp: excelapp, source: wb1, destination: output, worksheetname: "Sheet8", filter: territory, filterstartcell: 2, filtercolumn: 2, lastcolumn: "P", straightcopy: false);

                //copies static worksheets to new workbook
                foreach (string ws in staticTM)
                {
                    wb1.Worksheets[ws].Copy(output.Worksheets[1]);
                }

                //expands outline levels (data -> group) and copies worksheet, then collapses outline levels
                wb1.Worksheets["Sheet9"].Outline.ShowLevels(0, 3);

                copypaste(excelapp: excelapp, source: wb1, destination: output, worksheetname: "Sheet9", filter: territory, filterstartcell: 5, filtercolumn: 4, lastcolumn: "CN", straightcopy: true);

                wb1.Worksheets["Sheet9"].Outline.ShowLevels(0, 1);
                output.Worksheets["Sheet9"].Outline.ShowLevels(0, 1);

                copypaste(excelapp: excelapp, source: wb1, destination: output, worksheetname: "Sheet10", filter: territory, filterstartcell: 1, filtercolumn: 1, lastcolumn: "D", straightcopy: true);

                //deletes blank worksheet created when new workbook is created, copies table of contents and selects the upper left cell so that workbook opens to that sheet
                output.Worksheets["Sheet1"].Delete();
                wb1.Worksheets["Table of Contents"].Copy(output.Worksheets[1]);
                output.Worksheets["Table of Contents"].Range["B2"].Select();

                //saves workbook
                string path = @"C:\Users\User\Documents\TM\";
                string workbookpath = path + territory + "-TM - Q1 21 IC Payout Scorecard " + cleanDate + ".xlsx";

                //breaks link to workbook copied from
                output.BreakLink(Q121IC, XlLinkType.xlLinkTypeExcelLinks);

                //saves new excel sheet if file does not exist
                if (!File.Exists(workbookpath))
                {
                    output.SaveAs(workbookpath);
                }
                object misValue = System.Reflection.Missing.Value;

                //closes workbooks and quits excel
                wb1.Close(false, misValue, misValue);
                output.Close(false, misValue, misValue);
                excelapp.Quit();
            }

            //applies process used for sales territory managers to regional managers, with some exceptions as managers receive a different report
            foreach (KeyValuePair<string,string> entry in rsdRegion)
            {
                Excel.Application excelapp = new Excel.Application();
                excelapp.Visible = false;

                excelapp.DisplayAlerts = false;

                Excel.Workbook wb1 = excelapp.Workbooks.Open(Q121IC);

                Excel.Workbook output = excelapp.Workbooks.Add();

                //copies worksheets based on managers sales team
                if (entry.Key != "Specialty")
                {
                    foreach (string ws in staticPCPTM)
                    {
                        wb1.Worksheets[ws].Copy(output.Worksheets[1]);
                    }
                }

                if (entry.Key == "Specialty")
                {
                    wb1.Worksheets["Sheet7"].Copy(output.Worksheets[1]);
                }

                copypaste(excelapp: excelapp, source: wb1, destination: output, worksheetname: "Sheet8", filter: entry.Key, filterstartcell: 2, filtercolumn: 1, lastcolumn: "P", straightcopy: false);

                //copies non filtered excel sheets
                foreach (string ws in staticRSD)
                {
                    wb1.Worksheets[ws].Copy(output.Worksheets[1]);
                }

                //expands outline levels (data -> group) and copies worksheet, then collapses outline levels
                wb1.Worksheets["RSD IC Scorecard"].Outline.ShowLevels(0, 3);

                copypaste(excelapp: excelapp, source: wb1, destination: output, worksheetname: "Sheet11", filter: entry.Key, filterstartcell: 5, filtercolumn: 3, lastcolumn: "CN", straightcopy: true);

                wb1.Worksheets["RSD IC Scorecard"].Outline.ShowLevels(0, 1);
                output.Worksheets["RSD IC Scorecard"].Outline.ShowLevels(0, 1);

                copypaste(excelapp: excelapp, source: wb1, destination: output, worksheetname: "Sheet10", filter: entry.Value, filterstartcell: 1, filtercolumn: 1, lastcolumn: "D", straightcopy: true);

                //same save and closeout process as TM loop
                output.Worksheets["Sheet1"].Delete();
                wb1.Worksheets["Table of Contents"].Copy(output.Worksheets[1]);
                output.Worksheets["Table of Contents"].Range["B2"].Select();

                string path = @"C:\Users\User\Documents\RSD\";
                string workbookpath = path + entry.Key + "-RSD - Q1 21 IC Payout Scorecard " + cleanDate + ".xlsx";

                output.BreakLink(Q121IC, XlLinkType.xlLinkTypeExcelLinks);

                if (!File.Exists(workbookpath))
                {
                    output.SaveAs(workbookpath);
                }
                object misValue = System.Reflection.Missing.Value;

                wb1.Close(false, misValue, misValue);
                output.Close(false, misValue, misValue);
                excelapp.Quit();
            }
        }
    }
}
