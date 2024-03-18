using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Controls;
using System.Threading;
using System.Windows.Threading;
using System.Windows;
using System.Reflection;
using System.IO.Pipes;
using MyGoogleServices;
using Microsoft.Win32;

namespace Multiple_Export_Data
{
    public class TableExporter
    {
        private static TableExporter instance = null;
        private static readonly object padlock = new object();

        TableExporter()
        {
            settings = Settings.Instance;

            DispatcherTimer timer = new DispatcherTimer();

            timer.Interval = TimeSpan.FromMilliseconds(1000);
            timer.Tick += new EventHandler((object sender, EventArgs e) => { OnAddLog_TableExport?.Invoke(logQueue); logQueue.Clear(); });
            timer.Start();

            foreach (var item in settings.GetProgramSetting().GetServerTypes())
            {
                Dic_ServerType.Add(item.Value, item.ID);
            }

            foreach (var item in settings.GetProgramSetting().GetLegionTypes())
            {
                Dic_LegionType.Add(item.Value, item.ID);
            }
        }

        public static TableExporter Instance
        {
            get
            {
                lock (padlock)
                {
                    if (instance == null)
                    {
                        instance = new TableExporter();
                    }
                    return instance;
                }
            }
        }

        private readonly Settings settings;
        public void OpenSettings()
        {
            settings.OpenSettings();
        }

        public void ExportTable()
        {
            List<ExportSetting> exportList = new List<ExportSetting>();
        }

        #region TableExport

        private Dictionary<string, int> Dic_ServerType = new Dictionary<string, int>();
        private Dictionary<string, int> Dic_LegionType = new Dictionary<string, int>();

        public System.Action<Queue<string>> OnAddLog_TableExport;

        private Queue<string> logQueue = new Queue<string>();
        private static readonly object m_lockObj = new object();
        private void AddLog(string log)
        {

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            const string FORMAT = "[{0}] {1}";
            lock (m_lockObj)
            {
                logQueue.Enqueue(string.Format(FORMAT, timestamp, log));
            }
        }

        private void AddLog(string format, params object[] args)
        {
            AddLog(string.Format(format, args));
        }

        private void ExportTableFromWorksheet_Thread(object param)
        {
            // 파라미터 값이 Node 클래스가 아니면 종료
            if (param.GetType() != typeof(Node))
            {
                return;
            }
            // Node 타입으로 강제 캐스트(자료형이 Object 타입)
            var temp = (Node)param;

            ExportTableFromWorksheet(temp.key as ProgramSetting.TotalSheetSettingsInfo, temp.value as List<ExportSetting>);

            AddLog(string.Format("Thread {0} Close", temp.index));
            temp.Set();
        }

        /// <summary>
        /// 워크시트에서 지정한 형식대로 데이터 추출
        /// </summary>
        /// <param name="Key"></param>
        /// <param name="Value"></param>
        private void ExportTableFromWorksheet(ProgramSetting.TotalSheetSettingsInfo Key, List<ExportSetting> Value)
        {
            switch(Key.Type)
            {
                case "Excel":
                    AddLog("Excel Open");
                    //엑셀 실행
                    Excel.Application _excelApp = ExcelManager.OpenExcel(false);
                    _excelApp.DisplayAlerts = false;

                    Excel.Workbooks workbooks = _excelApp.Workbooks;

                    //파일 오픈
                    if (!File.Exists(Key.Path))
                    {
                        AddLog($"Excel WorkBook {Key.Name} Not Found, Path : Key.Path");
                        return;
                    }

                    AddLog($"Excel WorkBook {Key.Name} Open");
                    Excel.Workbook wb = workbooks.Open(Key.Path);

                    foreach (var temp in Value)
                    {
                        AddLog("출력 준비중 {0} {1} {2} {3} {4} ", Key.Name, temp.ServerType, temp.LegionType, temp.SourceName, temp.Path);
                        ExcelManager.ExportData_Async(_excelApp, wb, temp);
                        AddLog("{0} 출력 완료", temp.SourceName);
                    }

                    //엑셀 종료
                    wb.Close();
                    AddLog("Excel WorkBook {0} Closed", Key.Name);

                    _excelApp.DisplayAlerts = true;
                    _excelApp.Quit();
                    //사용한 워크북 및 프로세스 해제
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp);

                    AddLog($"Excel Closed");
                    return;
                case "Google":

                    AddLog("Excel Open");
                    //엑셀 실행
                    Excel.Application _excelApp1 = ExcelManager.OpenExcel(false);
                    _excelApp1.DisplayAlerts = false;
                    //파일 오픈

                    AddLog($"Excel WorkBook {Key.Name} Open");

                    RegistryKey reg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Funigloo.ExcelAddIn", false);
                    //IGoogleSheetManager _GSManager = ExcelManager.OpenGoogleSheet(reg.GetValue("CredentialPath").ToString());
                    IGoogleSheetManager _GSManager = ExcelManager.OpenGoogleSheet("credentials.json");
                    if (_GSManager == null)
                    {
                        AddLog("Google Sheet Manager Open Failed");
                        return;
                    }

                    foreach (var temp in Value)
                    {
                        AddLog("출력 준비중 {0} {1} {2} {3} {4} ", Key.Name, temp.ServerType, temp.LegionType, temp.SourceName, temp.Path);
                        ExcelManager.GS_ExportData_Async(_excelApp1, _GSManager,temp);
                        AddLog("{0} 출력 완료", temp.SourceName);
                    }

                    _excelApp1.DisplayAlerts = true;
                    _excelApp1.Quit();
                    //사용한 워크북 및 프로세스 해제
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp1);

                    AddLog($"Excel Closed");
                    break;
                default:
                    AddLog("잘못된 타입입니다.");
                    break;
            }
        }

        /// <summary>
        /// 선택된 출력 목록에 해당되는 워크시트와 출력 목록을 결합.
        /// </summary>
        /// <param name="totlaSheetsInfo"></param>
        /// <param name="exportSettings"></param>
        private Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>> SortExportSettings(List<ProgramSetting.TotalSheetSettingsInfo> totlaSheetsInfo, List<ExportSetting> exportSettings)
        {
            var result = new Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>>();

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
                            AddLog("이미 추가된 대상입니다.");
                            continue;
                        }
                        //추가
                        result[temp].Add(exportSetting);
                        AddLog($"출력 예정 리스트 {exportSetting.Name} 추가 완료");
                    }
                }
            }

            return result;
        }
        
        #region ExportTable 기본

        /// <summary>
        /// 선택한 출력 목록에 맞게 데이터 추출 후 출력(비동기)(반복문)
        /// </summary>
        /// <param name="data"></param>
        public async void ExportAll(ItemCollection data)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            var exportSettings = new List<ExportSetting>();
            Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>> Dic_Sort_By_TotalSheet = new Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>>();
            for (int index = 0; index < data.Count; ++index)
            {
                if (data[index] is System.Data.DataRowView)
                {
                    var temp = data[index] as System.Data.DataRowView;
                    if (temp.Row["IsSelected"] is bool && (bool)temp.Row["IsSelected"])
                    {
                        string serverType = temp.Row["SERVER_TYPE"].ToString();
                        string legionType = temp.Row["LEGION_TYPE"].ToString();
                        string sourceName = temp.Row["SOURCE_NAME"].ToString();

                        if (!Dic_ServerType.ContainsKey(serverType) || !Dic_LegionType.ContainsKey(legionType))
                        {
                            AddLog("잘못된 서버 타입 또는 지역 타입입니다.");
                            continue;
                        }
                        exportSettings.Add(settings.GetExportSetting(Dic_ServerType[serverType], Dic_LegionType[legionType], sourceName));
                    }
                }
            }
            Dic_Sort_By_TotalSheet = SortExportSettings(settings.GetProgramSetting().GetTotalSheetSettings(), exportSettings);

            await ExportTable(Dic_Sort_By_TotalSheet);

            sw.Stop();
            AddLog($" ExporAll : {sw.ElapsedMilliseconds}ms");
        }

        /// <summary>
        /// Task 전달용 ExportTable_Multiple
        /// </summary>
        /// <param name="Dic_Sort_By_TotalSheet"></param>
        private Task ExportTable(Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>> Dic_Sort_By_TotalSheet)
        {
            return Task.Factory.StartNew(() => ExportTable_Multiple(Dic_Sort_By_TotalSheet));
        }

        /// <summary>
        /// 데이터 출력 로직 (멀티스레드 사용 안함)
        /// </summary>
        /// <param name="Dic_Sort_By_TotalSheet"></param>
        private void ExportTable_Multiple(Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>> Dic_Sort_By_TotalSheet)
        {
            foreach (var item in Dic_Sort_By_TotalSheet)
            {
                if (item.Value.Count == 0)
                {
                    continue;
                }

                ExportTableFromWorksheet(item.Key, item.Value);
            }

            //엑셀 프로세스를 많이 사용했기때문에 확실하게 처리하기 위해 가비지 콜랙터 2번 사용
            //가비지 컬렉터 실행
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion ExportTable 기본

        #region ExportTable 멀티스레딩

        public class Node : EventWaitHandle
        {
            public Node() : base(false, EventResetMode.ManualReset)
            {

            }
            public int index { get; set; }
            public ProgramSetting.TotalSheetSettingsInfo key { get; set; }
            public List<ExportSetting> value { get; set; }
        }


        private EventWaitHandle AddNode(List<EventWaitHandle> list, int index, ProgramSetting.TotalSheetSettingsInfo key, List<ExportSetting> value)
        {
            var node = new Node { index = index, key = key, value = value };
            list.Add(node);
            return node;
        }

        /// <summary>
        /// 데이터 출력 로직. 멀티스레드
        /// </summary>
        /// <param name="Dic_Sort_By_TotalSheet"></param>
        private void ExportTable_Multiple_Async(Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>> Dic_Sort_By_TotalSheet)
        {
            int index = -1;
            var list = new List<EventWaitHandle>();

            foreach (var item in Dic_Sort_By_TotalSheet)
            {
                if(item.Value == null)
                {
                    continue;
                }
                else if (item.Value.Count == 0)
                {
                    continue;
                }
                try
                {
                    ThreadPool.QueueUserWorkItem(ExportTableFromWorksheet_Thread, AddNode(list,++index, item.Key, item.Value));
                    AddLog(string.Format("Thread {0} Start", index));
                }
                catch (Exception e)
                {
                    AddLog(e.Message);
                }
            }

            WaitHandle.WaitAll(list.ToArray());
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// 선택한 출력 목록에 맞게 데이터 추출 후 출력(비동기)(멀티스레딩)
        /// </summary>
        /// <param name="data"></param>
        public async void ExportAll_Async(ItemCollection data)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            var exportSettings = new List<ExportSetting>();
            Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>> Dic_Sort_By_TotalSheet = new Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>>();
            for (int index = 0; index < data.Count; ++index)
            {
                if (data[index] is System.Data.DataRowView)
                {
                    var temp = data[index] as System.Data.DataRowView;
                    if (temp.Row["IsSelected"] is bool && (bool)temp.Row["IsSelected"])
                    {
                        string serverType = temp.Row["SERVER_TYPE"].ToString();
                        string legionType = temp.Row["LEGION_TYPE"].ToString();
                        string sourceName = temp.Row["SOURCE_NAME"].ToString();

                        if (!Dic_ServerType.ContainsKey(serverType) || !Dic_LegionType.ContainsKey(legionType))
                        {
                            AddLog("잘못된 서버 타입 또는 지역 타입입니다.");
                            continue;
                        }
                        exportSettings.Add(settings.GetExportSetting(Dic_ServerType[serverType], Dic_LegionType[legionType], sourceName));
                    }
                }
            }
            Dic_Sort_By_TotalSheet = SortExportSettings(settings.GetProgramSetting().GetTotalSheetSettings(), exportSettings);

            AddLog($" ExporAll_Async : Target Sheet Sorted");

            try
            {
                await ExportTable_Async(Dic_Sort_By_TotalSheet);
            }
            catch(Exception e)
            {
                AddLog($" ExporAll_Async Error : {e.Message}");
            }

            sw.Stop();
            AddLog($" ExporAll_Async : {sw.ElapsedMilliseconds}ms");
            MessageBox.Show("출력 완료");
        }

        /// <summary>
        /// Task 전달용 ExportTable_Multiple_Async
        /// </summary>
        /// <param name="Dic_Sort_By_TotalSheet"></param>
        private Task ExportTable_Async(Dictionary<ProgramSetting.TotalSheetSettingsInfo, List<ExportSetting>> Dic_Sort_By_TotalSheet)
        {
            return Task.Factory.StartNew(() => ExportTable_Multiple_Async(Dic_Sort_By_TotalSheet));
        }
        #endregion

        #endregion
        
    }
}
