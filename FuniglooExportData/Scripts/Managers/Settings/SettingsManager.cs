using System.Collections.Generic;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuniglooExportData
{
    //singleton class
    public class SettingsManager
    {
        public enum PROGRASS_TYPE
        {
            None = 0,
            checkSettingsFileExist =  1,
            closeSettingsFile =  2,
            OpenSettingsFile =  3,
            InitializeProgramSetting =  4,
            InitializeProgramSettings =  5,
            InitializeExportSettings =  6,
        }

        public const string SETTING_FILE_NAME = "Settings.xlsx";
        private const string PROGRAM_SETTINGS_SHEET_NAME = "ProgramSettings";
        private const string EXPORT_SETTINGS_SHEET_NAME = "ExportSettings";
        private Excel.Application _settingsApplication;

        private Dictionary<string, Excel.Worksheet> _Dic_SettingSheets;
        private static SettingsManager instance = null;
        private static readonly object padlock = new object();

        public int TEST = 0;

        public System.Action<string> OnMessage;
        private string settingFilePath;

        public Dictionary<(int ServerType, int LegionType, string SourceName), ExportSetting> DIc_ExportSettings = new Dictionary<(int ServerType, int LegionType, string SourceName), ExportSetting>();
        public List<ExportSetting> List_ExportSettings = new List<ExportSetting>();
        private ProgramSetting programSetting;
        private bool isInitialized = false;

        public System.Action<ExportDataProcess.PROGRASS_TYPE> OnPrograssInit;

        public string CurrentWorksheetName { get; set; }

        public bool IsInitialized => isInitialized;

        public ProgramSetting ProgramSetting
        {
            get
            {
                return programSetting;
            }
        }

        SettingsManager()
        {

        }

        public static SettingsManager Instance
        {
            get
            {
                lock (padlock)
                {
                    if (instance == null)
                    {
                        instance = new SettingsManager();
                    }
                    return instance;
                }
            }
        }

        /// <summary>
        /// 초기화
        /// </summary>
        public void Initialize(string Path)
        {
            settingFilePath = Path;
            try
            {

                if (_settingsApplication == null)
                    _settingsApplication = new Excel.Application();
                if (Path != null)
                {
                    OnPrograssInit?.Invoke(ExportDataProcess.PROGRASS_TYPE.checkSettingsFileExist);
                    //filePath 파일 존재 여부 확인
                    if (File.Exists(Path))
                    {
                        //세팅 프로세스 시트 열려있는거 종료
                        if (_settingsApplication.ActiveWorkbook != null)
                            _settingsApplication.Workbooks.Close();

                        OnPrograssInit?.Invoke(ExportDataProcess.PROGRASS_TYPE.OpenSettingsFile);
                        //파일이 존재하면 열기
                        _settingsApplication.Workbooks.Open(Path);

                        OverWrapValues(_settingsApplication.ActiveWorkbook);
                        _settingsApplication.Workbooks.Close();
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
                OnPrograssInit?.Invoke(ExportDataProcess.PROGRASS_TYPE.closeSettingsFile);
                ShutdownSettingApplication();
                isInitialized = true;
            }
        }

        /// <summary>
        /// 세팅 앱 종료
        /// </summary>
        public void OnShutdownSettingApplication(object sender = null, System.EventArgs e = null)
        {
            ShutdownSettingApplication();
        }

        private void ShutdownSettingApplication()
        {
            if (_settingsApplication != null)
            {
                _settingsApplication.Workbooks.Close();
                _settingsApplication.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_settingsApplication);
                _settingsApplication = null;
            }
        }

        /// <summary>
        /// 세팅파엘에서 값 덮어쓰기
        /// </summary>
        void OverWrapValues(Excel.Workbook _settingWorkBook)
        {
            OnPrograssInit?.Invoke(ExportDataProcess.PROGRASS_TYPE.InitializeProgramSettings);
            if (_settingsApplication == null)
            {
                OnMessage?.Invoke("설정 파일 경로가 설정되지 않았거나 변경점이 있습니다. 재접속 후 확인 부탁드립니다.");
                return;
            }

            if (_Dic_SettingSheets == null)
            {
                _Dic_SettingSheets = new Dictionary<string, Excel.Worksheet>();
            }
            else
            {
                _Dic_SettingSheets.Clear();
            }

            foreach (var item in _settingWorkBook.Worksheets)
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

            OnPrograssInit?.Invoke(ExportDataProcess.PROGRASS_TYPE.InitializeProgramSetting);
            if (_Dic_SettingSheets.ContainsKey(PROGRAM_SETTINGS_SHEET_NAME))
            {
                var sht = _Dic_SettingSheets[PROGRAM_SETTINGS_SHEET_NAME];
                programSetting = new ProgramSetting(sht);
            }

            OnPrograssInit?.Invoke(ExportDataProcess.PROGRASS_TYPE.InitializeExportSettings);
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

        public ExportSetting GetExportSettingByName(string SourceName)
        {
            if (List_ExportSettings.Find(x => x.Name == SourceName) != null)
            {
                return List_ExportSettings.Find(x => x.Name == SourceName);
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
            if (_settingsApplication != null)
            {
                _settingsApplication.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(SettingsWorkbookBeforeClose);
                _settingsApplication.Workbooks.Open(settingFilePath, false, false);
                _settingsApplication.Visible = true;
            }
            else
            {
                _settingsApplication = new Excel.Application();
                _settingsApplication.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(SettingsWorkbookBeforeClose);
                _settingsApplication.Workbooks.Open(settingFilePath, false, false);
                _settingsApplication.Visible = true;
            }
        }

        void SettingsWorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
            //프로세스 종료 및 설정파일 재설정.
            OverWrapValues(wb);
        }
    }
}