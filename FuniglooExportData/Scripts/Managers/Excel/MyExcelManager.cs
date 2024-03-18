using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuniglooExportData
{
    // make singleton class
    public class MyExcelManager
    {
        public enum ERROR_CODE
        {
            SUCCESS = 0,
            ERROR = 1,
            OUT_OF_RANGE = 2,
            COPY_ERROR = 3,
        }
        public const string PATH_FORMAT = "{0}{1}_{2}\\";
        public const string LOG_EXPORT = "[Export] ";
        private SettingsManager settingsManager;

        private string latestExportDirectoryPath = "";
        private string latestExportFilePath;
        private List<string> ExportLogList = new List<string>();

        public List<string> ExportLogs => ExportLogList;

        public Action<string> OnExportFinish;


        public string LatestExportDirectoryPath => latestExportDirectoryPath;
        public string LatestExportFilePath => latestExportFilePath;

        private static MyExcelManager instance = null;
        private static readonly object padlock = new object();

        MyExcelManager()
        {
            settingsManager = SettingsManager.Instance;
        }

        public static MyExcelManager Instance
        {
            get
            {
                lock (padlock)
                {
                    if (instance == null)
                    {
                        instance = new MyExcelManager();
                    }
                    return instance;
                }
            }
        }


        public Excel.Workbook MakeNewWorkBook()
        {
            var wkbk = Globals.ThisAddIn.Application.Workbooks.Add();
            return wkbk;
        }

        public int GetOpenedSheetIndex()
        {
            var item = Globals.ThisAddIn.Application.ActiveSheet;

            if (item is Excel.Worksheet)
            {
                var sheet = item as Excel.Worksheet;
                return sheet.Index;
            }

            return 0;
        }

        public int GetIndexByName(Excel.Workbook thisworkbook, String name)
        {
            var sheets = thisworkbook.Worksheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                if (sheets[i].Name == name)
                {
                    return i;
                }
            }

            return 0;
        }

        private (string check, int check_index, int error_row_a, int error_row_b) error_check(Excel.Workbook thisworkbook, int targetIndex)
        {
            Excel.Worksheet tw;
            //var temp_value;

            string check = "없음";
            int check_index = 0;
            int error_row_a = 0;
            int error_row_b = 0;

            tw = thisworkbook.Worksheets[targetIndex];
            var check_range = tw.Range["A2", tw.Range["A2"].End[Excel.XlDirection.xlDown]];
            for (int i = 2; i < check_range.Count; ++i)
            {
                if (string.IsNullOrEmpty(tw.Cells[i, "BA"].Value))
                {
                    check_index = 1;
                }

                if (tw.Cells[i, "BA"] != null)
                {
                    check_index = 2;
                }

                if (check_index > 0)
                {
                    check = tw.Name;
                    error_row_a = i;
                    return (check, check_index, error_row_a, error_row_b);
                }
            }

            return (check, check_index, error_row_a, error_row_b);
        }

        public void ExportData(Excel.Workbook thisworkbook, int targetIndex)
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
            var newWorkBook = MakeNewWorkBook();

            if (newWorkBook.Worksheets.Count <= 0)
            {
                newWorkBook.Worksheets.Add();
            }

            Excel.Worksheet wksheet = sheets[targetIndex];
            var fileName = sheetName + ".xlsx";

            //var temp = wkbk.Worksheets.Count;
            var wkSheet1 = newWorkBook.Worksheets.Item[1];
            wksheet.Copy(wkSheet1);
            wkSheet1 = newWorkBook.Worksheets.Item[wksheet.Name];
            wkSheet1.Name = "Origin";

            var wkSheet2 = newWorkBook.Worksheets.Item["Sheet1"];

            if (CopyRange(wkSheet1, wkSheet2, "A:C,G:AY", "A1") == 1)
                return;

            newWorkBook.Worksheets.Item["Origin"].Delete();
            wkSheet2.Name = sheetName;
            newWorkBook.SaveAs(fileName);
            newWorkBook.Close(true);
        }

        public int _ExportData(Excel.Workbook thisworkbook, ExportSetting exportSetting)
        {
            //시트 확인
            var activeSheet = thisworkbook.ActiveSheet;

            if (!(activeSheet is Excel.Worksheet))
            {
                //ExportLogList.Enqueue("활성화된 시트가 없습니다.");
                return (int)ERROR_CODE.ERROR;
            }

            //통합문서 생성
            var newWorkbook = MakeNewWorkBook();

            if (newWorkbook.Worksheets.Count <= 0)
            {
                newWorkbook.Worksheets.Add();
            }

            //시트명 생성
            var sheetName = exportSetting.Name;
            //시트 경로 생성
            Excel.Worksheet originalWkSheet = activeSheet as Excel.Worksheet;
            string serverStr = settingsManager.ProgramSetting.GetServerType(exportSetting.ServerType).Value;
            string legionrStr = settingsManager.ProgramSetting.GetLegionType(exportSetting.ServerType).Value;
            var path = string.Format(settingsManager.ProgramSetting.PathFormat.Value.ToString, exportSetting.Path, serverStr, legionrStr);
            //시트경로 + 시트명 합산
            var fileName = string.Format("{0}{1}{2}", path, sheetName, settingsManager.ProgramSetting.DefaultExtension.Value.ToString);

            //생성한 통합 문서로 데이터 복사
            var copiedWkSheet = newWorkbook.Worksheets.Item[1];
            originalWkSheet.Copy(copiedWkSheet);
            copiedWkSheet = newWorkbook.Worksheets.Item[originalWkSheet.Name];
            copiedWkSheet.Name = "Origin";

            //복사한 데이터를 원하는 위치로 이동 및 적용
            var targetWkSheet = newWorkbook.Worksheets.Item["Sheet1"];
            int result = CopyRange(copiedWkSheet, targetWkSheet, exportSetting.SourceRange, exportSetting.TargetRange);
            //복사 실패시
            if (result != 0)
            {
                return (int)ERROR_CODE.COPY_ERROR;
            }
            //복사한 데이터 삭제 및 시트명 변경
            newWorkbook.Worksheets.Item["Origin"].Delete();

            //시트명 변경
            targetWkSheet.Name = sheetName;

            //통합문서 저장 및 종료
            newWorkbook.SaveAs(fileName);
            newWorkbook.Close(true);


            MessageBox.Show($"출력이 완료되었습니다. 파일 위치{fileName}");
            latestExportDirectoryPath = path;
            latestExportFilePath = fileName;
            ExportLogList.Add(fileName);
            OnExportFinish?.Invoke(LOG_EXPORT + fileName);
            return (int)ERROR_CODE.SUCCESS;

        }

        public int ExportData(Excel.Workbook thisworkbook, ExportSetting exportSetting)
        {
            //시트 내 exportSetting combine 검증
            byte _Check = 0;
            foreach (Excel.Worksheet sht in thisworkbook.Worksheets)
            {
                foreach (var range in exportSetting.CombinRanges)
                {
                    if (sht.Name == range.SheetName)
                    {
                        _Check++;
                    }
                }
            }
            if (_Check != exportSetting.CombineCount)
            {
                MessageBox.Show("CombineRange 값에 포함된 시트를 찾을 수 없습니다.");
                return (int)ERROR_CODE.ERROR;
            }

            //통합문서 생성
            var newWorkbook = MakeNewWorkBook();

            if (newWorkbook.Worksheets.Count <= 0)
            {
                newWorkbook.Worksheets.Add();
            }

            //시트명 생성
            var sheetName = exportSetting.Name;
            //시트 경로 생성
            string serverStr = settingsManager.ProgramSetting.GetServerType(exportSetting.ServerType).Value;
            string legionrStr = settingsManager.ProgramSetting.GetLegionType(exportSetting.LegionType).Value;
            var path = string.Format(settingsManager.ProgramSetting.PathFormat.Value.ToString, exportSetting.Path, serverStr, legionrStr);
            //시트경로 + 시트명 합산
            var fileName = path + sheetName + settingsManager.ProgramSetting.DefaultExtension.Value.ToString;

            Excel.Worksheet originalWkSheet;
            Excel.Worksheet copiedWkSheet;
            Excel.Worksheet targetWkSheet;
            string targetRange = exportSetting.TargetRange;
            int tempRowCount = 1;
            if (exportSetting.CombineCount != 0)
            {
                //복사한 데이터를 원하는 위치로 이동 및 적용
                targetWkSheet = newWorkbook.Worksheets.Item["Sheet1"];

                for (int i = 0; i < exportSetting.CombineCount; i++)
                {
                    //생성한 통합 문서로 데이터 복사
                    originalWkSheet = thisworkbook.Worksheets.Item[exportSetting.CombinRanges[i].SheetName] as Excel.Worksheet;
                    newWorkbook.Worksheets.Add();
                    copiedWkSheet = newWorkbook.Worksheets.Item["Sheet" + (i + 2)];
                    if (CopyRange(originalWkSheet, copiedWkSheet, exportSetting.CombinRanges[i].Range, exportSetting.TargetRange) != 0)
                        return (int)ERROR_CODE.COPY_ERROR;
                    int rowCount = GetSheetRowCount(copiedWkSheet);
                    //함수 많은경우 복사할때 겁나 오래 걸림
                    string temp = EditRange(exportSetting.SourceRange, i == 0 ? 1 : 2, rowCount);
                    //copiedWkSheet = newWorkbook.Worksheets.Item["Sheet2"];
                    copiedWkSheet.Name = "Origin" + i;

                    if (CopyRange(copiedWkSheet, targetWkSheet, temp, string.Format("A{0}", tempRowCount)) != 0)
                        return (int)ERROR_CODE.COPY_ERROR;

                    tempRowCount += rowCount - 1;
                    //복사한 데이터 삭제 및 시트명 변경
                    newWorkbook.Worksheets.Item["Origin" + i].Delete();
                }

                //시트명 변경
                targetWkSheet.Name = sheetName;
            }
            else
            {
                //생성한 통합 문서로 데이터 복사
                originalWkSheet = thisworkbook.ActiveSheet as Excel.Worksheet;
                copiedWkSheet = newWorkbook.Worksheets.Item[1];
                if (originalWkSheet == null)
                    MessageBox.Show("originalWkSheet == null.");
                if (copiedWkSheet == null)
                    MessageBox.Show("copiedWkSheet == null.");
                originalWkSheet.Copy(copiedWkSheet);
                copiedWkSheet = newWorkbook.Worksheets.Item[originalWkSheet.Name];
                copiedWkSheet.Name = "Origin";

                //복사한 데이터를 원하는 위치로 이동 및 적용
                targetWkSheet = newWorkbook.Worksheets.Item["Sheet1"];
                int result = CopyRange(copiedWkSheet, targetWkSheet, exportSetting.SourceRange, exportSetting.TargetRange);
                //복사 실패시
                if (result != 0)
                    return (int)ERROR_CODE.COPY_ERROR;

                //복사한 데이터 삭제 및 시트명 변경
                newWorkbook.Worksheets.Item["Origin"].Delete();
                //시트명 변경
                targetWkSheet.Name = sheetName;
            }

            if (!File.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //통합문서 저장 및 종료
            newWorkbook.SaveAs(fileName);
            newWorkbook.Close(true);


            MessageBox.Show($"출력이 완료되었습니다. 파일 위치{fileName}");
            latestExportDirectoryPath = path;
            latestExportFilePath = fileName;
            ExportLogList.Add(fileName);
            OnExportFinish?.Invoke(LOG_EXPORT + fileName);
            return (int)ERROR_CODE.SUCCESS;
        }

        //A function that copies the entered range of the worksheet entered as the first parameter and copies the value to the entered location of the worksheet entered as the second parameter
        private int CopyRange(Excel.Worksheet sourceWorksheet, Excel.Worksheet targetWorksheet, string sourceRange, string targetRange)
        {
            try
            {
                Excel.Range source = sourceWorksheet.Range[sourceRange];
                Excel.Range target = targetWorksheet.Range[targetRange];
                source.Copy();
                target.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                return 0;
            }
            catch (Exception ex)
            {
                var temp = ex.Message;
                MessageBox.Show(ex.Message);
                return 1;
            }
        }

        /// <summary>
        /// 시트 첫줄만 확인해서 Null값이 나오면 그 전까지의 RowCount를 반환
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private int GetSheetRowCount(Excel.Worksheet worksheet)
        {
            Excel.Range columms = worksheet.Columns["A"];

            List<CellData> cells = CellData.RangeToCellDataList(columms, true);
            int count = 0;
            foreach (var cell in cells)
            {
                if (cell.Value == null)
                    break;
                count++;
            }
            return count;
        }

        private string EditRange(string origin, int startRow, int endRow)
        {
            string[] strings = origin.Split(':');
            return string.Format("{0}:{1}", strings[0] + startRow, strings[1] + endRow);
        }
    }
}
