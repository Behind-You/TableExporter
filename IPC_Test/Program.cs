using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Pipes;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Net.Sockets;
using System.Security.AccessControl;
using System.Runtime.CompilerServices;
using System.Xml.Schema;
using static IPC_Test.ProgramSetting;
using Microsoft.Office.Interop.Excel;
using System.IO.Ports;

namespace IPC_Test
{
    public class Program
    {
        private static Server server;
        private static Client client;
        private static Settings settings => Settings.Instance;
        private static Excel.Application excelApp;
        private static bool IsExporting = false;

        private static void Initialize()
        {
            Settings.Instance.OnChangeStatus += ShowSettingsStatusLog;
            Settings.Instance.OnMessage += ShowMessage;

            ExcelManager.OnMessage += ShowExcelMessage;
            ExcelManager.OnExportFinish += ShowExcelMessage;
            ExcelManager.OnLog += ShowExcelMessage;

            RegistryKey reg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Funigloo.ExcelAddIn", false);
            Settings.Instance.Initialize(reg.GetValue("SettingPath").ToString());
        }

        static void Main(string[] args)
        {
            Initialize();

            while (true)
            {
                if (IsExporting)
                    continue;

                Console.WriteLine("Select Mode");
                Console.WriteLine("1. Server");
                Console.WriteLine("2. Client");
                Console.WriteLine("3. Open Settings");
                Console.WriteLine("5. Open TotalSheet List");
                Console.WriteLine("6. ExportData");
                Console.WriteLine("7. Validate TotalSheet List");
                Console.WriteLine("8. Export TotalSheet");
                Console.WriteLine("99. Exit");
                string answer = Console.ReadLine();

                switch (answer)
                {
                    case "1":
                        {
                            Console.WriteLine("Server Mode");
                            server = new Server();
                            string port = settings.GetProgramSetting().GetProgramSetting(3).Value.ToString;
                            string ip = settings.GetProgramSetting().GetProgramSetting(4).Value.ToString;
                            server.StartSocket(port, ip);
                            break;
                        }
                    case "2":
                        {
                            Console.WriteLine("Client Mode");
                            client = new Client();
                            string port = settings.GetProgramSetting().GetProgramSetting(3).Value.ToString;
                            string ip = settings.GetProgramSetting().GetProgramSetting(4).Value.ToString;
                            client.StartSocketClient(port, ip);
                            break;
                        }
                    case "3":
                        {
                            settings.OpenSettings();
                            break;
                        }
                    case "4":
                        {
                            Console.WriteLine("Excel Open");
                            break;
                        }
                    case "5":
                        {
                            Console.WriteLine("Excel Open");
                            var totlaSheetsInfo = settings.GetProgramSetting().GetTotalSheetSettings();
                            int index = 0;
                            foreach (var item in totlaSheetsInfo)
                            {
                                Console.WriteLine(string.Format("{0}. {1}", index++, item.Name));
                            }
                        }
                        break;
                    case "6":
                        {
                            Console.WriteLine("Excel Open");
                            var totlaSheetsInfo = settings.GetProgramSetting().GetTotalSheetSettings();
                            int index = 0;
                            //관리 시트 리스트 출력
                            foreach (var item in totlaSheetsInfo)
                            {
                                Console.WriteLine(string.Format("{0}. {1}", index++, item.Name));
                            }

                            //관리 시트 선택
                            int result_index = 0;
                            ProgramSetting.TotalSheetSettingsInfo selectedSettingsInfo;
                            while (true)
                            {
                                Console.WriteLine("Select Number");
                                string result = Console.ReadLine();

                                if (result == "Stop")
                                    return;


                                if (!int.TryParse(result, out result_index))
                                {
                                    Console.WriteLine("숫자가 아닙니다. 다시 입력해주세요");
                                    continue;
                                }

                                selectedSettingsInfo = totlaSheetsInfo[result_index];
                                break;
                            }

                            excelApp = ExcelManager.OpenExcel();
                            Excel.Workbooks workbooks = excelApp.Workbooks;

                            if (!File.Exists(selectedSettingsInfo.Path))
                                continue;
                            Excel.Workbook wb = workbooks.Open(selectedSettingsInfo.Path);
                            ValidateTotalSheets(wb, selectedSettingsInfo);

                            if (excelApp != null)
                            {
                                wb.Close();
                                excelApp.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                                excelApp = null;
                            }
                            //가비지 컬렉터 실행
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        break;
                    case "7":
                        {
                            Console.WriteLine("Excel Open");
                            var totlaSheetsInfo = settings.GetProgramSetting().GetTotalSheetSettings();
                            int index = 0;

                            //관리 시트 리스트 출력
                            foreach (var item in totlaSheetsInfo)
                            {
                                Console.WriteLine(string.Format("{0}. {1}", index++, item.Name));

                                //엑셀 실행
                                excelApp = ExcelManager.OpenExcel(false);

                                Excel.Workbooks workbooks = excelApp.Workbooks;

                                //파일 오픈
                                if (!File.Exists(item.Path))
                                    continue;
                                Excel.Workbook wb = workbooks.Open(item.Path);

                                //시트 검증
                                int falsecount = ValidateTotalSheets(wb, item);

                                //결과 출력
                                Console.WriteLine("({0}/{1})", falsecount, item.ContainSources.Count);
                                Console.WriteLine();

                                //엑셀 종료
                                wb.Close();
                                excelApp.Quit();
                                //사용한 워크북 및 프로세스 해제
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                                excelApp = null;
                            }

                            //엑셀 프로세스를 많이 사용했기때문에 확실하게 처리하기 위해 가비지 콜랙터 2번 사용
                            //가비지 컬렉터 실행
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            //가비지 컬렉터 실행
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        break;
                    case "8":
                        {
                            Console.WriteLine("Excel Open");
                            var programSetting = settings.GetProgramSetting();
                            var totlaSheetsInfo = programSetting.GetTotalSheetSettings();
                            var serverList = programSetting.GetServerTypes();
                            var legionList = programSetting.GetLegionTypes();
                            int index = 0;
                            int serverType = 0;
                            int legionType = 0;
                            List<ExportSetting> exportList = new List<ExportSetting>();
                            Dictionary<TotalSheetSettingsInfo, List<ExportSetting>> Dic_Sort_By_TotalSheet = new Dictionary<TotalSheetSettingsInfo, List<ExportSetting>>();
                            string ans;
                            //입력 영역
                            while (true)
                            {
                                Console.WriteLine("출력할 서버를 선택해주세요");
                                index = 0;
                                foreach (var item in serverList)
                                {
                                    Console.WriteLine(string.Format("{0}. {1}", index++, item.Name));
                                }
                                ans = Console.ReadLine();

                                if (int.TryParse(ans, out serverType))
                                    break;

                                Console.WriteLine("입력값이 숫자가 아닙니다.");
                            }

                            while (true)
                            {
                                Console.WriteLine("출력할 지역을 선택해주세요");
                                index = 0;
                                foreach (var item in legionList)
                                {
                                    Console.WriteLine(string.Format("{0}. {1}", index++, item.Name));
                                }
                                ans = Console.ReadLine();

                                if (int.TryParse(ans, out legionType))
                                    break;

                                Console.WriteLine("입력값이 숫자가 아닙니다.");
                            }

                            var exportSettings = settings.GetExportSettings(serverType, legionType);
                            while (true)
                            {
                                Console.WriteLine("출력할 대상을 선택해주세요");
                                index = 0;
                                foreach (var item in exportSettings)
                                {
                                    Console.WriteLine(string.Format("{0}. {1}", index++, item.Name));
                                }

                                if (int.TryParse(Console.ReadLine(), out int target))
                                {
                                    if (target > exportSettings.Count)
                                    {
                                        Console.WriteLine("입력값이 리스트 범위를 벗어났습니다.");
                                        continue;
                                    }

                                    var exportSetting = exportSettings[target];
                                    /*
                                    foreach (var temp in totlaSheetsInfo)
                                    {
                                        //관리 시트에 대상이 포함되어있는지 확인
                                        if (temp.ContainSources.Contains(exportSetting.SourceName))
                                        {
                                            //키가 없으면 추가
                                            if (!Dic_Sort_By_TotalSheet.ContainsKey(temp))
                                            {
                                                Dic_Sort_By_TotalSheet.Add(temp, new List<ExportSetting>());
                                            }
                                            //이미 추가된 대상인지 확인
                                            if (Dic_Sort_By_TotalSheet[temp].Contains(exportSetting))
                                            {
                                                Console.WriteLine("이미 추가된 대상입니다.");
                                                break;
                                            }
                                            //추가
                                            Dic_Sort_By_TotalSheet[temp].Add(exportSetting);
                                            break;
                                        }
                                    }
                                    */
                                    exportList.Add(exportSetting);


                                    Console.WriteLine("추가로 입력하시겠습니까? (Y/N)");
                                    if(IsContinue(Console.ReadLine()))
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("입력값이 숫자가 아닙니다.");
                                }
                            }

                            Dic_Sort_By_TotalSheet = SortExportSettings(totlaSheetsInfo, exportList);

                            ExportMultipleTable(Dic_Sort_By_TotalSheet);
                        }
                        break;
                    case "10":
                        {
                            Console.WriteLine("Excel Open");
                            excelApp = ExcelManager.OpenExcel();

                            if (!File.Exists("E:\\기획\\Table\\A_라큐관리테이블_v3.xlsm"))
                                continue;
                            Excel.Workbooks workBooks = excelApp.Workbooks;
                            Excel.Workbook wb = workBooks.Open("E:\\기획\\Table\\A_라큐관리테이블_v3.xlsm");

                            //알람 끄기
                            excelApp.DisplayAlerts = false;
                            int result_index = 0;
                            while (true)
                            {
                                Console.WriteLine("몇번 세팅으로 출력할까요?");
                                string result = Console.ReadLine();

                                if (!int.TryParse(result, out result_index))
                                {
                                    Console.WriteLine("5숫자가 아닙니다. 다시 입력해주세요");
                                    continue;
                                }
                                break;
                            }

                            ExcelManager.ExportData(excelApp, wb, Settings.Instance.GetExportSetting(result_index));
                            //알람 켜기
                            excelApp.DisplayAlerts = true;

                            wb.Close();
                            excelApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBooks);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                            //가비지 컬렉터 실행
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        break;
                    case "99":
                        {
                            Console.WriteLine("Exit");
                            return;
                        }
                }
            }            
        }
        private static bool IsContinue(string ans)
        {
            switch (ans)
            {
                case "YES":
                case "Yes":
                case "yes":
                case "TRUE":
                case "True":
                case "true":
                case "1":
                case "Y":
                case "y":
                    {
                        return true;
                    };

                case "N":
                case "n":
                    {
                        return false;
                    }

                default:
                    {
                        return false;
                    }
            }
        }

        private static int ValidateTotalSheets(Excel.Workbook wb, ProgramSetting.TotalSheetSettingsInfo selectedSettingsInfo)
        {
            int falseCount = 0;
            Dictionary<string, bool> ValidDic = new Dictionary<string, bool>();

            try
            {
                foreach (string item in selectedSettingsInfo.ContainSources)
                {
                    ValidDic.Add(item, false);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("겹치는 항목 존재. Settings 파일의 TotalSheetContainsSourceNames 확인");
                return 0;
            }

            try
            {
                foreach (Excel.Worksheet item in wb.Sheets)
                {
                    if (ValidDic.ContainsKey(item.Name))
                        ValidDic[item.Name] = true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("WorkBook 관련 문제. 워크북 할당 확인");
                return 0;
            }

            foreach (var item in ValidDic)
            {
                Console.WriteLine(string.Format("{0} : {1}", item.Key, item.Value));
                if (item.Value == false)
                    falseCount++;
            }
            return falseCount;
        }

        private static void ExportMultipleTable(Dictionary<TotalSheetSettingsInfo, List<ExportSetting>> Dic_Sort_By_TotalSheet)
        {
            Console.WriteLine("Excel Open");
            int index = 0;

            Console.WriteLine("출력 예정 리스트.");
            foreach (var item in Dic_Sort_By_TotalSheet)
            {
                foreach (var temp in item.Value)
                {
                    Console.WriteLine(string.Format("{0} {1} {2} {3} {4} {5}", index++, item.Key.Name, temp.ServerType, temp.LegionType, temp.SourceName, temp.Path));
                }
            }

            index = 0;
            foreach (var item in Dic_Sort_By_TotalSheet)
            {
                //엑셀 실행
                Excel.Application _excelApp = ExcelManager.OpenExcel(false);
                _excelApp.DisplayAlerts = false;

                Excel.Workbooks workbooks = _excelApp.Workbooks;

                //파일 오픈
                if (!File.Exists(item.Key.Path))
                    continue;
                Excel.Workbook wb = workbooks.Open(item.Key.Path);

                foreach (var temp in item.Value)
                {
                    Console.WriteLine(string.Format("출력 준비중 {0} {1} {2} {3} {4} {5}", index++, item.Key.Name, temp.ServerType, temp.LegionType, temp.SourceName, temp.Path));
                    ExcelManager.ExportData(_excelApp, wb, temp);
                    Console.WriteLine("출력 완료");
                }

                //엑셀 종료
                wb.Close();

                _excelApp.DisplayAlerts = true;
                _excelApp.Quit();
                //사용한 워크북 및 프로세스 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp);
            }

            //엑셀 프로세스를 많이 사용했기때문에 확실하게 처리하기 위해 가비지 콜랙터 2번 사용
            //가비지 컬렉터 실행
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static Dictionary<TotalSheetSettingsInfo, List<ExportSetting>> SortExportSettings(List<TotalSheetSettingsInfo> totlaSheetsInfo, List<ExportSetting> exportSettings)
        {
            var result = new Dictionary<TotalSheetSettingsInfo, List<ExportSetting>>();

            foreach (var temp in totlaSheetsInfo)
            {
                //키가 없으면 추가
                if (!result.ContainsKey(temp))
                {
                    result.Add(temp, new List<ExportSetting>());
                }

                foreach (var exportSetting in exportSettings)
                {
                    //관리 시트에 대상이 포함되어있는지 확인
                    if (temp.ContainSources.Contains(exportSetting.SourceName))
                    {
                        //이미 추가된 대상인지 확인
                        if (result[temp].Contains(exportSetting))
                        {
                            Console.WriteLine("이미 추가된 대상입니다.");
                            continue;
                        }
                        //추가
                        result[temp].Add(exportSetting);
                    }
                }
            }

            return result;
        }

        private static void ExportData(Excel.Workbook Wb)
        {
            if (IsExporting)
            {
                Console.WriteLine("이미 실행중입니다.");
                return;
            }
            IsExporting = true;

            IsExporting = false;
        }

        static void ShowMessage(string message)
        {
            Console.WriteLine(message);
        }

        static void ShowSettingsStatusLog(Settings.Current_Status status)
        {
            Console.WriteLine("[Settings] " + status.ToString());
        }

        static void ShowExcelMessage(string message)
        {
            Console.WriteLine("[Excel] " + message);
        }
    }
    public class Settings
    {
        private static Settings instance = null;
        private static readonly object padlock = new object();

        Settings()
        {

        }

        public static Settings Instance
        {
            get
            {
                lock (padlock)
                {
                    if (instance == null)
                    {
                        instance = new Settings();
                    }
                    return instance;
                }
            }
        }

        public enum Current_Status
        {
            NotInitialized = 0,
            checkSettingsPath,
            checkSettingsFileExist,
            OpenSettingsFile,
            LoadSettingsFile,
            InitializeProgramSetting,
            InitializeExportSettings,
            closeSettingsFile,
            Ready = 99,
        }

        public const string SETTING_FILE_NAME = "Settings.xlsx";
        private const string PROGRAM_SETTINGS_SHEET_NAME = "ProgramSettings";
        private const string EXPORT_SETTINGS_SHEET_NAME = "ExportSettings";

        private Dictionary<string, Excel.Worksheet> _Dic_SettingSheets;

        public System.Action<string> OnMessage;
        public System.Action<Current_Status> OnChangeStatus;
        private string settingFilePath;

        private Dictionary<(int ServerType, int LegionType), List<ExportSetting>> _Cash_Dic_ExportSettings_Server_Legion = new Dictionary<(int ServerType, int LegionType), List<ExportSetting>>();
        private Dictionary<(int LegionType, string SourceName), List<ExportSetting>> _Cash_Dic_ExportSettings_Legion_SourceName = new Dictionary<(int LegionType, string SourceName), List<ExportSetting>>();

        public Dictionary<(int ServerType, int LegionType, string SourceName), ExportSetting> DIc_ExportSettings = new Dictionary<(int ServerType, int LegionType, string SourceName), ExportSetting>();
        public List<ExportSetting> List_ExportSettings = new List<ExportSetting>();
        private ProgramSetting programSetting;
        private bool isInitialized = false;

        public string CurrentWorksheetName { get; set; }

        public bool IsInitialized => isInitialized;


        /// <summary>
        /// 초기화
        /// </summary>
        public void Initialize(string Path)
        {
            settingFilePath = Path;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            try
            {
                OnChangeStatus?.Invoke(Current_Status.checkSettingsPath);
                if (Path != null)
                {
                    //filePath 파일 존재 여부 확인
                    OnChangeStatus?.Invoke(Current_Status.checkSettingsFileExist);
                    if (File.Exists(Path))
                    {
                        OnChangeStatus?.Invoke(Current_Status.OpenSettingsFile);
                        //파일이 존재하면 열기
                        workbooks.Open(Path);
                        Excel.Workbook wb = excelApp.ActiveWorkbook;
                        Excel.Sheets wkShts = wb.Worksheets;
                        OverWrapValues(wkShts);

                        workbooks.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wkShts);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    }
                    else
                    {
                        OnMessage?.Invoke("설정 파일 경로가 설정되지 않았습니다.");
                    }
                }
                else
                {
                    OnMessage?.Invoke("설정 파일 경로가 설정되지 않았습니다.");
                }
            }
            finally
            {
                if (excelApp != null)
                {
                    workbooks.Close();
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                    GC.Collect();
                }
                isInitialized = true;
            }
        }


        /// <summary>
        /// 설정파일에서 값 불러와 덮어쓰기
        /// </summary>
        void OverWrapValues(Excel.Sheets wkSheets)
        {
            if (_Dic_SettingSheets == null)
            {
                _Dic_SettingSheets = new Dictionary<string, Excel.Worksheet>();
            }
            else
            {
                _Dic_SettingSheets.Clear();
            }

            OnChangeStatus?.Invoke(Current_Status.LoadSettingsFile);
            foreach (var item in wkSheets)
            {
                if (item is Excel.Worksheet)
                {
                    var sht = item as Excel.Worksheet;

                    if (_Dic_SettingSheets.ContainsKey(sht.Name))
                    {
                        _Dic_SettingSheets[sht.Name] = sht;
                    }
                    else
                    {
                        _Dic_SettingSheets.Add(sht.Name, sht);
                    }
                }
            }

            OnChangeStatus?.Invoke(Current_Status.InitializeProgramSetting);
            if (_Dic_SettingSheets.ContainsKey(PROGRAM_SETTINGS_SHEET_NAME))
            {
                var sht = _Dic_SettingSheets[PROGRAM_SETTINGS_SHEET_NAME];
                programSetting = new ProgramSetting(sht);
            }

            OnChangeStatus?.Invoke(Current_Status.InitializeExportSettings);
            if (_Dic_SettingSheets.ContainsKey(EXPORT_SETTINGS_SHEET_NAME))
            {
                var sht = _Dic_SettingSheets[EXPORT_SETTINGS_SHEET_NAME];
                var exportSettings = ExportSetting.GetExportSettings(sht, programSetting.ExportSettingRange.Value.ToString);
                List_ExportSettings = exportSettings.List;
                DIc_ExportSettings = exportSettings.Dic;
            }
        }

        public int GetExportSettingsCount()
        {
            return List_ExportSettings.Count;
        }


        /// <summary>
        /// ExportSetting 가져오기
        /// </summary>
        public ExportSetting GetExportSetting(int index)
        {
            if (List_ExportSettings.Count > index)
            {
                return List_ExportSettings[index];
            }
            return null;
        }

        /// <summary>
        /// ExportSetting 가져오기
        /// </summary>
        public ExportSetting GetExportSetting(int ServerType, int LegionType, string SourceName)
        {
            if (DIc_ExportSettings.ContainsKey((ServerType, LegionType, SourceName)))
            {
                return DIc_ExportSettings[(ServerType, LegionType, SourceName)];
            }
            return null;
        }

        public ExportSetting GetExportSettingsByName(string SourceName)
        {
            if (List_ExportSettings.Find(x => x.Name == SourceName) != null)
            {
                return List_ExportSettings.Find(x => x.Name == SourceName);
            }
            return null;
        }

        public ExportSetting FindExportSettings(Predicate<ExportSetting> match)
        {
            if (List_ExportSettings.Find(match) != null)
            {
                return List_ExportSettings.Find(match);
            }
            return null;
        }

        public bool IsExportSettingExist(int ServerType, int LegionType, string SourceName)
        {
            if (DIc_ExportSettings.ContainsKey((ServerType, LegionType, SourceName)))
            {
                return true;
            }
            return false;
        }

        public bool IsExportSettingExist(int index)
        {
            if (List_ExportSettings.Count > index)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Settings 파일 열기
        /// </summary>
        public void OpenSettings()
        {
            if (settingFilePath == null)
            {
                OnMessage?.Invoke("설정 파일 경로가 설정되지 않았습니다.");
                return;
            }
            else
            {
                var _settingsApplication = new Excel.Application();
                _settingsApplication.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(SettingsFileBeforeClose);
                _settingsApplication.Workbooks.Open(settingFilePath, false, false);
                _settingsApplication.Visible = true;
            }
        }

        void SettingsFileBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
            Excel.Sheets wksheets = wb.Worksheets;
            //프로세스 종료 및 설정파일 재설정.
            OverWrapValues(wksheets);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(wksheets);
        }

        public ProgramSetting GetProgramSetting()
        {
            return programSetting;
        }

        public string GetServerTypeValue(int ServerType)
        {
            return programSetting.GetServerType(ServerType).Value;
        }

        public string GetLegionTypeValue(int LegionType)
        {
            return programSetting.GetLegionType(LegionType).Value;
        }

        public string GetExcelPathFormatValue()
        {
            return programSetting.PathFormat.Value.ToString;
        }
        public string GetExcelDefaultExtension()
        {
            return programSetting.DefaultExtension.Value.ToString;
        }

        public List<ExportSetting> GetExportSettings()
        {
            return List_ExportSettings;
        }

        public List<ExportSetting> GetExportSettings(int _serverType, int _legionType)
        {
            //캐시에 있으면 전달
            if (_Cash_Dic_ExportSettings_Server_Legion.ContainsKey((_serverType, _legionType)))
            {
                return _Cash_Dic_ExportSettings_Server_Legion[(_serverType, _legionType)];
            }
            else
            {
                //없으면 생성해서 전달
                if (List_ExportSettings.FindAll(x => x.ServerType == _serverType && x.LegionType == _legionType) != null)
                {
                    _Cash_Dic_ExportSettings_Server_Legion.Add((_serverType, _legionType), List_ExportSettings.FindAll(x => x.ServerType == _serverType && x.LegionType == _legionType));
                    return _Cash_Dic_ExportSettings_Server_Legion[(_serverType, _legionType)];
                }
                else
                {
                    //아에 없으면 null
                    return null;
                }
            }
        }

    }

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
    }



    public class Server
    {
        public Server()
        {

        }

        public void StartPipe()
        {
            //파이프 클라이언트 생성
            Process pipeClient = new Process();

            //파이프 클라이언트 명칭 설정.
            pipeClient.StartInfo.FileName = "pipeClient.exe";

            using (AnonymousPipeServerStream pipeServer =
                new AnonymousPipeServerStream(PipeDirection.Out,
                HandleInheritability.Inheritable))
            {
                Console.WriteLine("[SERVER] Current TransmissionMode: {0}.",
                    pipeServer.TransmissionMode);

                // Pass the client process a handle to the server.
                pipeClient.StartInfo.Arguments =
                    pipeServer.GetClientHandleAsString();
                pipeClient.StartInfo.UseShellExecute = false;
                pipeClient.Start();

                pipeServer.DisposeLocalCopyOfClientHandle();

                try
                {
                    // Read user input and send that to the client process.
                    using (StreamWriter sw = new StreamWriter(pipeServer))
                    {
                        sw.AutoFlush = true;
                        // Send a 'sync message' and wait for client to receive it.
                        sw.WriteLine("SYNC");
                        pipeServer.WaitForPipeDrain();
                        // Send the console input to the client process.
                        Console.Write("[SERVER] Enter text: ");
                        sw.WriteLine(Console.ReadLine());
                    }
                }
                // Catch the IOException that is raised if the pipe is broken
                // or disconnected.
                catch (IOException e)
                {
                    Console.WriteLine("[SERVER] Error: {0}", e.Message);
                }
            }

            pipeClient.WaitForExit();
            pipeClient.Close();
            Console.WriteLine("[SERVER] Client quit. Server terminating.");
        }


        public void StartSocket(string port, string ip)
        {
            Console.WriteLine("Server Start");
            Console.WriteLine($"Port Number : {port}");
            Console.WriteLine($"IP Address : {ip}");

            System.Net.IPAddress ipAddr = System.Net.IPAddress.Parse(ip);
            System.Net.IPEndPoint ipEndPoint = new System.Net.IPEndPoint(ipAddr, int.Parse(port));

            System.Net.Sockets.Socket server = new System.Net.Sockets.Socket(System.Net.Sockets.AddressFamily.InterNetwork, System.Net.Sockets.SocketType.Stream, System.Net.Sockets.ProtocolType.Tcp);

            server.Bind(ipEndPoint);
            for (int i = 0; i < 10; ++i)
            {
                server.Listen(10);
                Console.WriteLine("Listening...");

                System.Net.Sockets.Socket client = server.Accept();

                client.Send(UTF8Encoding.UTF8.GetBytes($"Connect Success ({i})"));

                byte[] data = new byte[1024];
                int size = client.Receive(data);
                Protocol.HandleRecieve(data);

                client.Close();
            }


            server.Close();
        }


    }




    public class Client
    {
        public Client()
        {

        }

        public void StartSocketClient(string port, string ip)
        {
            Console.WriteLine("Client Start");
            Console.WriteLine($"Port Number : {port}");
            Console.WriteLine($"IP Address : {ip}");

            System.Net.IPAddress ipAddr = System.Net.IPAddress.Parse(ip);
            System.Net.IPEndPoint ipEndPoint = new System.Net.IPEndPoint(ipAddr, int.Parse(port));

            System.Net.Sockets.Socket client = new System.Net.Sockets.Socket(System.Net.Sockets.AddressFamily.InterNetwork, System.Net.Sockets.SocketType.Stream, System.Net.Sockets.ProtocolType.Tcp);

            client.Connect(ipEndPoint);

            byte[] data = new byte[1024];
            int size = client.Receive(data);

            Console.WriteLine(UTF8Encoding.UTF8.GetString(data));

            client.Send(Protocol.REQUEST_EXPORT_ACTIVATED_SHEET(0, 0));



            client.Close();
        }
    }

    public class Protocol
    {
        public enum ProtocolTypes
        {
            EXPORT_ACTIVATED_SHEET = 1,
        }

        /// <summary>
        /// 프로토콜 기본 바이트 사이즈. 프로토콜 넘버 + 파라미터 개수 = int + int = 4 + 4 = 8
        /// </summary>
        public const int PROTOCOL_BASIZ_BYTESIZE = sizeof(int) * 2;

        public static byte[] REQUEST_EXPORT_ACTIVATED_SHEET(int _ServerType, int _LegionType)
        {
            byte[] data = new byte[PROTOCOL_BASIZ_BYTESIZE + (sizeof(int) * 2)];
            int paramIndex = 0;
            //프로토콜 넘버
            BitConverter.GetBytes((int)ProtocolTypes.EXPORT_ACTIVATED_SHEET).CopyTo(data, sizeof(int) * paramIndex);
            paramIndex++;
            //파라미터 갯수
            BitConverter.GetBytes(2).CopyTo(data, sizeof(int) * paramIndex);
            paramIndex++;

            BitConverter.GetBytes(_ServerType).CopyTo(data, sizeof(int) * paramIndex);
            paramIndex++;
            BitConverter.GetBytes(_LegionType).CopyTo(data, sizeof(int) * paramIndex);

            return data;
        }

        public static void HandleRecieve(byte[] data)
        {
            int paramIndex = 0;
            Console.WriteLine(data);
            ProtocolTypes protocolType = (ProtocolTypes)BitConverter.ToInt32(data, sizeof(int) * paramIndex);
            switch (protocolType)
            {
                case ProtocolTypes.EXPORT_ACTIVATED_SHEET:
                    {
                        Console.WriteLine("EXPORT_ACTIVATED_SHEET Recieved");
                        paramIndex++;
                        Console.WriteLine($"ServerType {BitConverter.ToInt32(data, sizeof(int) * paramIndex)}");
                        paramIndex++;
                        Console.WriteLine($"LegionType {BitConverter.ToInt32(data, sizeof(int) * paramIndex)}");
                    }
                    break;
            }
        }

    }
}