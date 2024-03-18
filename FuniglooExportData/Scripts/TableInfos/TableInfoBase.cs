using Excel = Microsoft.Office.Interop.Excel;

namespace FuniglooExportData
{
    public class TableInfoBase
    {
        public readonly string Root;
        public readonly string Name;
        public readonly byte Type;
        public readonly string DefaultRange;
        public readonly string DefaultPasteLocation;

        public TableInfoBase(string root, string name, byte type, string defaultRange, string defaultPasteLocation)
        {
            Root = root;
            Name = name;
            Type = type;
            DefaultRange = defaultRange;
            DefaultPasteLocation = defaultPasteLocation;
        }

        public Excel.Workbook MakeNewWorkBook()
        {
            var wkbk = Globals.ThisAddIn.Application.Workbooks.Add();
            return wkbk;
        }

        //A function that copies the entered range of the worksheet entered as the first parameter and copies the value to the entered location of the worksheet entered as the second parameter
        private void CopyRange(Excel.Worksheet sourceWorksheet, Excel.Worksheet targetWorksheet, string sourceRange, string targetRange)
        {
            Excel.Range source = sourceWorksheet.Range[sourceRange];
            Excel.Range target = targetWorksheet.Range[targetRange];
            source.Copy();
            target.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, true, false);
        }



        public void MergeData(Excel.Workbook thisworkbooks)
        {

        }

        public virtual void ExportData(Excel.Workbook thisworkbook, int targetIndex)
        {
            if (targetIndex == 0)
            {
                return;
            }

            Excel.Sheets sheets = thisworkbook.Worksheets;
            if (sheets.Count <= 0 || sheets.Count <= targetIndex)
            {
                return;
            }

            var sheetName = sheets[targetIndex].Name;// + "_" + serverType.ToString();
            var wkbk = MakeNewWorkBook();

            if (wkbk.Worksheets.Count <= 0)
            {
                wkbk.Worksheets.Add();
            }

            Excel.Worksheet wksheet = sheets[targetIndex];
            var fileName = sheetName + ".xlsx";

            //var temp = wkbk.Worksheets.Count;
            var wkSheet1 = wkbk.Worksheets.Item[1];
            wksheet.Copy(wkSheet1);
            wkSheet1 = wkbk.Worksheets.Item[wksheet.Name];
            wkSheet1.Name = "Origin";

            var wkSheet2 = wkbk.Worksheets.Item["Sheet1"];

            if (CopyRange(wkSheet1, wkSheet2, "A:C,G:AY", "A1") == 1)
                return;

            wkbk.Worksheets.Item["Origin"].Delete();
            wkSheet2.Name = sheetName;
            wkbk.SaveAs(fileName);
            wkbk.Close(true);
        }
    }

}
