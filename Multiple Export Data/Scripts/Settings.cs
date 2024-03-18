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

namespace Multiple_Export_Data
{
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

                OnChangeStatus?.Invoke(Current_Status.closeSettingsFile);
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
                OpenWorkBook(settingFilePath);
            }
        }

        public void OpenWorkBook(string path)
        {
            if (path == null)
            {
                OnMessage?.Invoke("설정 파일 경로가 설정되지 않았습니다.");
                return;
            }
            else
            {
                //구글 시트 링크면 오픈
                if(path.Contains("https://"))
                {
                    System.Diagnostics.Process.Start(path);
                    return;
                }
                else if (path.Contains("xlsm") || path.Contains("xlsx") || path.Contains("xls") || path.Contains("csv") || path.Contains("xlsb"))
                {
                    if (!File.Exists(path))
                    {
                        OnMessage?.Invoke("파일이 존재하지 않습니다.");
                        return;
                    }
                    var _settingsApplication = new Excel.Application();
                    _settingsApplication.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(SettingsFileBeforeClose);
                    _settingsApplication.Workbooks.Open(path, false, false);
                    _settingsApplication.Visible = true;
                }
                else
                {
                    return;
                }
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

        public int GetServerTypeID(string val)
        {
           return programSetting.GetServerTypes().Find(x => x.Value == val).ID;
        }

        public int GetLegionTypeID(string val)
        {
            return programSetting.GetLegionTypes().Find(x => x.Value == val).ID;
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

        public System.Data.DataTable GetExportSettings_ToDataTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("IsSelected", typeof(bool));
            dt.Columns.Add(ExportSetting.ExportParameter.INDEX.ToString(), typeof(int));
            dt.Columns.Add(ExportSetting.ExportParameter.SERVER_TYPE.ToString(),    typeof(string));
            dt.Columns.Add(ExportSetting.ExportParameter.LEGION_TYPE.ToString(),    typeof(string));
            dt.Columns.Add(ExportSetting.ExportParameter.NAME.ToString(),           typeof(string));
            dt.Columns.Add(ExportSetting.ExportParameter.PATH.ToString(),           typeof(string));
            dt.Columns.Add(ExportSetting.ExportParameter.SOURCE_NAME.ToString(),    typeof(string));
            dt.Columns.Add(ExportSetting.ExportParameter.SOURCE_RANGE.ToString(),   typeof(string));
            dt.Columns.Add(ExportSetting.ExportParameter.TARGET_RANGE.ToString(),   typeof(string));
            dt.Columns.Add(ExportSetting.ExportParameter.COMBINE_COUNT.ToString(),  typeof(string));
            for(int i = 0; i < 20; ++i)
            {
                dt.Columns.Add(string.Format("COMBINE_RANGE_{0}", i), typeof(string));
            }

            foreach(var item in List_ExportSettings)
            {
                System.Data.DataRow dr = dt.NewRow();
                dr.SetField("IsSelected", false);
                dr.SetField(ExportSetting.ExportParameter.INDEX.ToString(),         item.Index);
                dr.SetField(ExportSetting.ExportParameter.SERVER_TYPE.ToString(),   GetServerTypeValue(item.ServerType));
                dr.SetField(ExportSetting.ExportParameter.LEGION_TYPE.ToString(),   GetLegionTypeValue(item.LegionType));
                dr.SetField(ExportSetting.ExportParameter.NAME.ToString(),          item.Name);
                dr.SetField(ExportSetting.ExportParameter.PATH.ToString(),          item.Path);
                dr.SetField(ExportSetting.ExportParameter.SOURCE_NAME.ToString(),   item.SourceName);
                dr.SetField(ExportSetting.ExportParameter.SOURCE_RANGE.ToString(),  item.SourceRange);
                dr.SetField(ExportSetting.ExportParameter.TARGET_RANGE.ToString(),  item.TargetRange);
                dr.SetField(ExportSetting.ExportParameter.COMBINE_COUNT.ToString(), item.CombineCount);
                for (int i = 0; i < item.CombinRanges.Count; ++i)
                {
                    dr.SetField(string.Format("COMBINE_RANGE_{0}", i), item.CombinRanges[i].Origin);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public System.Data.DataTable GetProgramSettings_ProgramSettings_ToDataTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("Value", typeof(string));

            foreach (var item in programSetting.GetProgramSettings())
            {
                System.Data.DataRow dr = dt.NewRow();
                dr.SetField("ID", item.ID);
                dr.SetField("Name", item.Name);
                dr.SetField("Type", item.Value.Type.Name);
                dr.SetField("Value", item.Value.ToString);
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public System.Data.DataTable GetProgramSettings_ServerTypes_ToDataTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("NAME", typeof(string));
            dt.Columns.Add("Value", typeof(string));
            dt.Columns.Add("ID", typeof(int));

            foreach (var item in programSetting.GetServerTypes())
            {
                System.Data.DataRow dr = dt.NewRow();
                dr.SetField("NAME", item.Name);
                dr.SetField("Value", item.Value);
                dr.SetField("ID", item.ID);
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public System.Data.DataTable GetProgramSettings_LegionTypes_ToDataTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("NAME", typeof(string));
            dt.Columns.Add("Value", typeof(string));
            dt.Columns.Add("ID", typeof(int));

            foreach (var item in programSetting.GetLegionTypes())
            {
                System.Data.DataRow dr = dt.NewRow();
                dr.SetField("NAME", item.Name);
                dr.SetField("Value", item.Value);
                dr.SetField("ID", item.ID);
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public System.Data.DataTable GetProgramSettings_TotalSheet_ToDataTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("NAME", typeof(string));
            dt.Columns.Add("PATH", typeof(string));
            dt.Columns.Add("SHEETS", typeof(string));

            foreach (var item in programSetting.GetTotalSheetSettings())
            {
                System.Data.DataRow dr = dt.NewRow();
                dr.SetField("NAME", item.Name);
                dr.SetField("PATH", item.Path);
                dr.SetField("SHEETS", item.ContainsSourceNames);
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public ObservableCollection<ExportSetting> GetExportSettings_ToObservableCollection()
        {
            ObservableCollection<ExportSetting> _observableCollection = new ObservableCollection<ExportSetting>();
            foreach (var item in List_ExportSettings)
            {
                _observableCollection.Add(item);
            }
            return _observableCollection;
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
}
