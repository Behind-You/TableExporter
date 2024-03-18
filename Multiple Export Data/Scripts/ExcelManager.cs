using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Pipes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Net.Sockets;
using System.Security.AccessControl;
using System.Runtime.CompilerServices;
using System.Xml.Schema;
using System.Data;
using System.Collections.ObjectModel;
using MyGoogleServices;
using Microsoft.Office.Interop.Excel;
using Google.Apis.Sheets.v4.Data;
using System.Windows.Controls;
using System.Xml.Linq;
using System.Windows.Input;

namespace Multiple_Export_Data
{
    public class ExcelManager
    {
        public const string PATH_FORMAT = "{0}{1}_{2}\\";
        public const string LOG_EXPORT = "[Export] ";
        const string LOG_FORMAT_ERROR = "[Error] {0}";


        private static List<string> ExportLogList = new List<string>();
        public static List<string> ExportLogs => ExportLogList;

        public static Action<string> OnExportFinish;
        public static Action<int> OnExportPrograss;
        public static System.Action<string> OnMessage;
        public static System.Action<string> OnLog;

        public static void Initialize()
        {
            ExportLogList.Clear();
        }

        static void AddLog(string log, bool isMessage = false)
        {
            ExportLogList.Add(log);
            OnLog?.Invoke(log);
            if (isMessage)
                OnMessage?.Invoke(log);
        }

        public static Excel.Application OpenExcel(bool isvisiable = true)
        {
            Excel.Application ExcelApp = new Excel.Application();
            ExcelApp.Visible = isvisiable;

            return ExcelApp;
        }

        public static IGoogleSheetManager OpenGoogleSheet()
        {
            GoogleConnectionMannager manager;
            manager = new GoogleConnectionMannager("credentials.json", new GoogleSheetManager());
            return manager.ActivatedSheetManager;
        }

        public static IGoogleSheetManager OpenGoogleSheet(string credentialsPath)
        {
            try
            {
                AddLog("Google Login Start");
                GoogleConnectionMannager manager;
                AddLog(String.Format("Google Login Credential Path : {0}", credentialsPath));
                manager = new GoogleConnectionMannager(credentialsPath, new GoogleSheetManager());
                return manager.ActivatedSheetManager;
            }
            catch (Exception ex)
            {
                AddLog(ex.Message, true);
                return null;
            }
        }

        public static Excel.Workbook MakeNewWorkBook(Excel.Application application)
        {
            var wkbk = application.Workbooks.Add();
            AddLog("WorkBook Added");
            return wkbk;
        }

        public static int GetOpenedSheetIndex(Excel.Application application)
        {
            var item = application.ActiveSheet;

            if (item is Excel.Worksheet)
            {
                var sheet = item as Excel.Worksheet;
                return sheet.Index;
            }

            return 0;
        }
        public static int GetIndexByName(Excel.Workbook thisworkbook, String name)
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

        /// <summary>
        /// 시트 첫줄만 확인해서 Null값이 나오면 그 전까지의 RowCount를 반환
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private static int GetSheetRowCount(Excel.Worksheet worksheet)
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

        private static string EditRange(string origin, int startRow, int endRow)
        {
            string[] strings = origin.Split(':');
            return string.Format("{0}:{1}", strings[0] + startRow, strings[1] + endRow);
        }

        public static void ExportData(Excel.Application application, Excel.Workbook thisworkbook, ExportSetting exportSetting)
        {
            OnExportPrograss?.Invoke(0);

            Excel.Sheets thisWorkSheets = thisworkbook.Worksheets;

            //시트 내 exportSetting combine 검증
            byte _Check = 0;
            foreach (Excel.Worksheet sht in thisWorkSheets)
            {
                foreach (var range in exportSetting.CombinRanges)
                {
                    if (sht.Name == range.SheetName)
                    {
                        _Check++;
                    }
                }
            }

            OnExportPrograss?.Invoke(1);
            if (_Check != exportSetting.CombineCount)
            {
                OnMessage?.Invoke("CombineRange 값에 포함된 시트를 찾을 수 없습니다.");
                OnLog(string.Format(LOG_FORMAT_ERROR, "CombineRange 값에 포함된 시트를 찾을 수 없습니다."));
            }

            OnExportPrograss?.Invoke(2);
            //통합문서 생성
            var newWorkbook = MakeNewWorkBook(application);
            var newWorkSheets = newWorkbook.Worksheets;

            if (newWorkSheets.Count <= 0)
            {
                newWorkSheets.Add();
            }

            OnExportPrograss?.Invoke(3);
            //시트명 생성
            var sheetName = exportSetting.Name;
            //시트 경로 생성
            string serverStr = Settings.Instance.GetServerTypeValue(exportSetting.ServerType);
            string legionrStr = Settings.Instance.GetLegionTypeValue(exportSetting.LegionType);
            var path = string.Format(Settings.Instance.GetExcelPathFormatValue(), exportSetting.Path, serverStr, legionrStr);
            //시트경로 + 시트명 합산
            var fileName = path + sheetName + Settings.Instance.GetExcelDefaultExtension();


            OnExportPrograss?.Invoke(4);
            Excel.Worksheet originalWkSheet;
            Excel.Worksheet copiedWkSheet;
            Excel.Worksheet targetWkSheet;
            string targetRange = exportSetting.TargetRange;
            int tempRowCount = 1;

            if (exportSetting.CombineCount != 0)
            {
                OnExportPrograss?.Invoke(5);
                //복사한 데이터를 원하는 위치로 이동 및 적용
                targetWkSheet = newWorkSheets.Item["Sheet1"];

                for (int i = 0; i < exportSetting.CombineCount; i++)
                {
                    //생성한 통합 문서로 데이터 복사
                    originalWkSheet = thisWorkSheets.Item[exportSetting.CombinRanges[i].SheetName] as Excel.Worksheet;
                    newWorkSheets.Add();
                    copiedWkSheet = newWorkSheets.Item["Sheet" + (i + 2)];
                    if (CopyRange(originalWkSheet, copiedWkSheet, exportSetting.CombinRanges[i].Range, exportSetting.TargetRange) != 0)
                    {
                        AddLog(string.Format("CopyRange Error Original {0} targer {1} OriginRange {2} TargetRange {3} ", originalWkSheet.Name, copiedWkSheet.Name, exportSetting.CombinRanges[i].Range, exportSetting.TargetRange), true);

                        newWorkbook.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWkSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkSheets);
                        return;
                    }
                    int rowCount = GetSheetRowCount(copiedWkSheet);
                    //함수 많은경우 복사할때 겁나 오래 걸림
                    string temp = EditRange(exportSetting.SourceRange, i == 0 ? 1 : 2, rowCount);
                    //copiedWkSheet = newWorkbook.Worksheets.Item["Sheet2"];
                    copiedWkSheet.Name = "Origin" + i;

                    if (CopyRange(copiedWkSheet, targetWkSheet, temp, string.Format("A{0}", tempRowCount)) != 0)
                    {
                        AddLog(string.Format("CopyRange Error Original {0} targer {1} OriginRange {2} TargetRange {3} ", copiedWkSheet.Name, targetWkSheet.Name, temp, string.Format("A{0}", tempRowCount)), true);
                        newWorkbook.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWkSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkSheets);
                        return;
                    }

                    tempRowCount += rowCount - 1;
                    //복사한 데이터 삭제 및 시트명 변경
                    newWorkSheets.Item["Origin" + i].Delete();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
                    OnExportPrograss?.Invoke(6 + i);
                }

                //시트명 변경
                targetWkSheet.Name = sheetName;
            }
            else
            {
                //생성한 통합 문서로 데이터 복사
                originalWkSheet = thisWorkSheets.Item[exportSetting.SourceName] as Excel.Worksheet;
                copiedWkSheet = newWorkSheets.Item[1];
                if (originalWkSheet == null)
                    AddLog("originalWkSheet == null.", true);
                if (copiedWkSheet == null)
                    AddLog("copiedWkSheet == null.", true);
                originalWkSheet.Copy(copiedWkSheet);

                copiedWkSheet = newWorkSheets.Item[originalWkSheet.Name];
                copiedWkSheet.Name = "Origin";

                //복사한 데이터를 원하는 위치로 이동 및 적용
                targetWkSheet = newWorkSheets.Item["Sheet1"];
                int result = CopyRange(copiedWkSheet, targetWkSheet, exportSetting.SourceRange, exportSetting.TargetRange);

                //복사 실패시
                if (result != 0)
                {
                    AddLog(string.Format("CopyRange Error Original {0} targer {1} OriginRange {2} TargetRange {3} ", copiedWkSheet.Name, targetWkSheet.Name, exportSetting.SourceRange, exportSetting.TargetRange), true);
                    application.Visible = true;
                    //newWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkSheets);
                    return;
                }

                //복사한 데이터 삭제 및 시트명 변경
                newWorkSheets.Item["Origin"].Delete();
                //시트명 변경
                targetWkSheet.Name = sheetName;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
            }


            if (!File.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //통합문서 저장 및 종료
            newWorkbook.SaveAs(fileName);
            newWorkbook.Close(true);


            AddLog($"출력이 완료되었습니다. 파일 위치{fileName}", true);
            AddLog(fileName);
            OnExportFinish?.Invoke(LOG_EXPORT + fileName);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkSheets);
            return;
        }

        public static async void ExportData_Async(Excel.Application application, Excel.Workbook thisworkbook, ExportSetting exportSetting)
        {
            OnExportPrograss?.Invoke(0);

            Excel.Sheets thisWorkSheets = thisworkbook.Worksheets;

            //시트 내 exportSetting combine 검증
            byte _Check = 0;
            foreach (Excel.Worksheet sht in thisWorkSheets)
            {
                foreach (var range in exportSetting.CombinRanges)
                {
                    if (sht.Name == range.SheetName)
                    {
                        _Check++;
                    }
                }
            }

            OnExportPrograss?.Invoke(1);
            if (_Check != exportSetting.CombineCount)
            {
                OnMessage?.Invoke("CombineRange 값에 포함된 시트를 찾을 수 없습니다.");
                OnLog(string.Format(LOG_FORMAT_ERROR, "CombineRange 값에 포함된 시트를 찾을 수 없습니다."));
            }

            OnExportPrograss?.Invoke(2);
            //통합문서 생성
            var newWorkbook = MakeNewWorkBook(application);
            var newWorkSheets = newWorkbook.Worksheets;

            if (newWorkSheets.Count <= 0)
            {
                newWorkSheets.Add();
            }

            OnExportPrograss?.Invoke(3);
            //시트명 생성
            var sheetName = exportSetting.Name;
            //시트 경로 생성
            string serverStr = Settings.Instance.GetServerTypeValue(exportSetting.ServerType);
            string legionrStr = Settings.Instance.GetLegionTypeValue(exportSetting.LegionType);
            var path = string.Format(Settings.Instance.GetExcelPathFormatValue(), exportSetting.Path, serverStr, legionrStr);
            //시트경로 + 시트명 합산
            var fileName = path + sheetName + Settings.Instance.GetExcelDefaultExtension();

            OnExportPrograss?.Invoke(4);
            Excel.Worksheet targetWkSheet;

            if (exportSetting.CombineCount != 0)
            {
                targetWkSheet = ExportWithCombine(newWorkSheets, thisWorkSheets, newWorkbook, exportSetting, application);
            }
            else
            {
                targetWkSheet = ExportOnly(newWorkSheets, thisWorkSheets, newWorkbook, exportSetting, application);
            }

            //시트명 변경
            targetWkSheet.Name = sheetName;


            if (!File.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //통합문서 저장 및 종료
            newWorkbook.SaveAs(fileName);
            newWorkbook.Close(true);


            AddLog($"출력이 완료되었습니다. 파일 위치{fileName}", true);
            AddLog(fileName);
            OnExportFinish?.Invoke(LOG_EXPORT + fileName);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkSheets);
            return;
        }

        public static async void GS_ExportData_Async(Excel.Application application, IGoogleSheetManager gsManager, ExportSetting exportSetting)
        {
            OnExportPrograss?.Invoke(0);

            OnExportPrograss?.Invoke(1);

            OnExportPrograss?.Invoke(2);
            //통합문서 생성
            Excel.Workbook newWorkbook = MakeNewWorkBook(application);

            OnExportPrograss?.Invoke(3);
            //시트명 생성
            var sheetName = exportSetting.Name;
            //시트 경로 생성
            string serverStr = Settings.Instance.GetServerTypeValue(exportSetting.ServerType);
            string legionrStr = Settings.Instance.GetLegionTypeValue(exportSetting.LegionType);
            var path = string.Format(Settings.Instance.GetExcelPathFormatValue(), exportSetting.Path, serverStr, legionrStr);
            //시트경로 + 시트명 합산
            var fileName = path + sheetName + Settings.Instance.GetExcelDefaultExtension();

            OnExportPrograss?.Invoke(4);

            //생성한 통합 문서로 데이터 복사
            //IList<IList<object>> originalData = gsManager.GetSheet(exportSetting.SourceName, exportSetting.SourceRange);

            var sourceRange = exportSetting.SourceRange.Split('!');
            if (sourceRange.Length != 2)
                return;
            var vRange = gsManager.GetSheetValueRanges(exportSetting.SourceName, sourceRange[0], sourceRange[1].Split(','));
            int result = GSCopyRange(newWorkbook, vRange, exportSetting.SourceRange,exportSetting.Name);

            //복사 실패시
            if (result != 0)
            {
                AddLog(string.Format("GSCopyRange Fail exportSetting.Name = {0}, exportSetting.SourceRange = {1}, exportSetting.TargetRange = {2} ", exportSetting.Name, exportSetting.SourceRange, exportSetting.TargetRange), true);
                application.Visible = true;
                //newWorkbook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
                return;
            }

            if (!File.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //통합문서 저장 및 종료
            newWorkbook.SaveAs(fileName);
            newWorkbook.Close(true);


            AddLog($"출력이 완료되었습니다. 파일 위치{fileName}", true);
            AddLog(fileName);
            OnExportFinish?.Invoke(LOG_EXPORT + fileName);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
            return;
        }

        private static Task<Excel.Worksheet> ExportWithCombine_Async(Excel.Sheets newWorkSheets, Excel.Sheets thisWorkSheets, Excel.Workbook newWorkbook, ExportSetting exportSetting, Excel.Application application)
        {
            return Task.Factory.StartNew(() => ExportWithCombine(newWorkSheets, thisWorkSheets, newWorkbook, exportSetting, application));
        }

        private static Task<Excel.Worksheet> ExportOnly_Async(Excel.Sheets newWorkSheets, Excel.Sheets thisWorkSheets, Excel.Workbook newWorkbook, ExportSetting exportSetting, Excel.Application application)
        {
            return Task.Factory.StartNew(() => ExportOnly(newWorkSheets, thisWorkSheets, newWorkbook, exportSetting, application));
        }

        private static Excel.Worksheet ExportWithCombine(Excel.Sheets newWorkSheets, Excel.Sheets thisWorkSheets, Excel.Workbook newWorkbook, ExportSetting exportSetting, Excel.Application application)
        {
            int tempRowCount = 1;
            OnExportPrograss?.Invoke(5);
            //복사한 데이터를 원하는 위치로 이동 및 적용
            Excel.Worksheet targetWkSheet = newWorkSheets.Item["Sheet1"];
            Excel.Worksheet originalWkSheet;
            Excel.Worksheet copiedWkSheet;
            for (int i = 0; i < exportSetting.CombineCount; i++)
            {
                //생성한 통합 문서로 데이터 복사
                originalWkSheet = thisWorkSheets.Item[exportSetting.CombinRanges[i].SheetName] as Excel.Worksheet;
                newWorkSheets.Add();
                copiedWkSheet = newWorkSheets.Item["Sheet" + (i + 2)];
                if (CopyRange(originalWkSheet, copiedWkSheet, exportSetting.CombinRanges[i].Range, exportSetting.TargetRange) != 0)
                {
                    AddLog(string.Format("CopyRange Error Original {0} targer {1} OriginRange {2} TargetRange {3} ", originalWkSheet.Name, copiedWkSheet.Name, exportSetting.CombinRanges[i].Range, exportSetting.TargetRange), true);

                    application.Visible = true;
                    //newWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkSheets);
                    return null;
                }
                int rowCount = GetSheetRowCount(copiedWkSheet);
                //함수 많은경우 복사할때 겁나 오래 걸림
                string temp = EditRange(exportSetting.SourceRange, i == 0 ? 1 : 2, rowCount);
                //copiedWkSheet = newWorkbook.Worksheets.Item["Sheet2"];
                copiedWkSheet.Name = "Origin" + i;

                if (CopyRange(copiedWkSheet, targetWkSheet, temp, string.Format("A{0}", tempRowCount)) != 0)
                {
                    AddLog(string.Format("CopyRange Error Original {0} targer {1} OriginRange {2} TargetRange {3} ", copiedWkSheet.Name, targetWkSheet.Name, temp, string.Format("A{0}", tempRowCount)), true);

                    application.Visible = true;
                    //newWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkSheets);
                    return null;
                }

                tempRowCount += rowCount - 1;
                //복사한 데이터 삭제 및 시트명 변경
                newWorkSheets.Item["Origin" + i].Delete();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
                OnExportPrograss?.Invoke(6 + i);
            }
            return targetWkSheet;
        }

        private static Excel.Worksheet ExportOnly(Excel.Sheets newWorkSheets, Excel.Sheets thisWorkSheets, Excel.Workbook newWorkbook, ExportSetting exportSetting, Excel.Application application)
        {
            //생성한 통합 문서로 데이터 복사
            Excel.Worksheet originalWkSheet = thisWorkSheets.Item[exportSetting.SourceName] as Excel.Worksheet;
            Excel.Worksheet copiedWkSheet = newWorkSheets.Item[1];
            if (originalWkSheet == null)
                AddLog("originalWkSheet == null.", true);
            if (copiedWkSheet == null)
                AddLog("copiedWkSheet == null.", true);
            originalWkSheet.Copy(copiedWkSheet);

            copiedWkSheet = newWorkSheets.Item[originalWkSheet.Name];
            copiedWkSheet.Name = "Origin";

            //복사한 데이터를 원하는 위치로 이동 및 적용
            Excel.Worksheet targetWkSheet = newWorkSheets.Item["Sheet1"];
            int result = CopyRange(copiedWkSheet, targetWkSheet, exportSetting.SourceRange, exportSetting.TargetRange);

            //복사 실패시
            if (result != 0)
            {
                AddLog(string.Format("CopyRange Error Original {0} targer {1} OriginRange {2} TargetRange {3} ", copiedWkSheet.Name, targetWkSheet.Name, exportSetting.SourceRange, exportSetting.TargetRange), true);
                application.Visible = true;
                //newWorkbook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorkSheets);
                return null;
            }


            System.Runtime.InteropServices.Marshal.ReleaseComObject(originalWkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(copiedWkSheet);
            //복사한 데이터 삭제 및 시트명 변경
            newWorkSheets.Item["Origin"].Delete();
            return targetWkSheet;
        }

        //A function that copies the entered range of the worksheet entered as the first parameter and copies the value to the entered location of the worksheet entered as the second parameter
        private static int CopyRange(Excel.Worksheet sourceWorksheet, Excel.Worksheet targetWorksheet, string sourceRange, string targetRange)
        {
            try
            {
                Excel.Range source = sourceWorksheet.Range[sourceRange];
                Excel.Range target = targetWorksheet.Range[targetRange];
                source.Copy();
                target.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(source);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(target);
                return 0;
            }
            catch (Exception ex)
            {
                var temp = ex.Message;
                AddLog(ex.Message, true);
                return 1;
            }
        }

        //A function that copies the entered range of the worksheet entered as the first parameter and copies the value to the entered location of the worksheet entered as the second parameter
        private static int GSCopyRange(Excel.Workbook targetWorkbook, ValueRange vRange, string range,string name)
        {
            try
            {
                Excel.Sheets targetWorksheets = targetWorkbook.Worksheets;
                Excel.Worksheet targetWorksheet = null;
                if (targetWorksheets.Count <= 0)
                {
                    targetWorksheets.Add();
                }
                else
                {
                    foreach (Excel.Worksheet item in targetWorksheets)
                    {
                        AddLog(item.Name, true);
                        targetWorksheet = item;
                    }
                }
                if (targetWorksheet == null)
                    return 1;

                IList<IList<object>> temp = vRange.Values;
                int colNum = 1;
                int rowNum = 1;
                foreach (var item in temp)
                {
                    foreach (var item2 in item)
                    {
                        string tempRange = ExcelRangeString(rowNum, colNum);
                        targetWorksheet.Range[tempRange].Value = item2;
                        rowNum++;
                    }
                    colNum++;
                    rowNum = 1;
                }
                // Excel.Range target = targetWorksheet.Range[range];
                // target.Value = vRange.Values;
                targetWorksheet.Name = name;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWorksheets);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWorksheet);
                return 0;
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                AddLog(message, true);
                return 1;
            }
        }
        private static int GSCopyRange(Excel.Workbook targetWorkbook, ValueRange[] vRange, string range, string name)
        {
            try
            {
                Excel.Sheets targetWorksheets = targetWorkbook.Worksheets;
                Excel.Worksheet targetWorksheet = null;
                if (targetWorksheets.Count <= 0)
                {
                    targetWorksheets.Add();
                }
                else
                {
                    foreach (Excel.Worksheet item in targetWorksheets)
                    {
                        AddLog(item.Name, true);
                        targetWorksheet = item;
                    }
                }
                if (targetWorksheet == null)
                    return 1;

                int colNum = 1;
                foreach (var v in vRange)
                {
                    IList<IList<object>> temp = v.Values;
                    GSCopyColumns(targetWorksheet,v,colNum, out colNum);
                }
                targetWorksheet.Name = name;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWorksheets);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(targetWorksheet);

                return 0;
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                AddLog(message, true);
                return 1;
            }
        }

        //컬럼 단위 데이터 쓰기
        private static int GSCopyColumns(Excel.Worksheet targetWorksheet, ValueRange vRange, int colNum, out int nextCol)
        {
            try
            {
                if (targetWorksheet == null)
                {
                    nextCol = colNum;
                    return 1;
                }

                IList<IList<object>> temp = vRange.Values;
                var colCount = temp[0].Count;
                int rowNum;
                for(int i = 0; i < colCount; ++i)
                {
                    rowNum = 1;
                    foreach (var item in temp)
                    {
                        string tempRange = ExcelRangeString(colNum + i, rowNum);
                        targetWorksheet.Range[tempRange].Value = item[i];
                        rowNum++;
                    }
                }
                nextCol = colNum + colCount;
                return 0;
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                AddLog(message, true);
                nextCol = colNum;
                return 1;
            }
        }

        private static Dictionary<int, string> _memo_ExcelColumnString = new Dictionary<int, string>();

        private static string ExcelRangeString(int col, int row)
        {
            string colStr = GetExcelColumnString(col);
            return colStr + row;
        }

        private static string GetExcelColumnString(int col)
        {
            if (!_memo_ExcelColumnString.ContainsKey(col))
            {
                string colStr = "";
                int key = col - 1;
                while (key >= 0)
                {
                    int remainder = key % 26;
                    colStr = (char)(remainder + 'A') + colStr;
                    key = key / 26 - 1;
                }
                _memo_ExcelColumnString[col] = colStr;
            }
            return _memo_ExcelColumnString[col];
        }

    }
}
