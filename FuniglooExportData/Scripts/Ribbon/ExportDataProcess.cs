using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuniglooExportData
{
    public class ExportDataProcess : MyRibbon.IPresenter, MyUserControl.IPresenter
    {
        #region 싱글톤 영역
        private static ExportDataProcess instance = null;
        private static readonly object padlock = new object();

        public static ExportDataProcess Instance
        {
            get
            {
                lock (padlock)
                {
                    if (instance == null)
                    {
                        instance = new ExportDataProcess();
                    }
                    return instance;
                }
            }
        }
        #endregion

        private const string MSGTITLE = "Funigloo Export Table Data";
        private const string MSG_NO_SETTINGS_FILE = "설정 파일을 찾을 수 없습니다.";
        private const string MSG_NO_SHEET = "선택한 시트가 없습니다.";
        private const string MSG_NO_DATA = "데이터가 없습니다.";
        private const string MSG_NO_EXPORT_TARGET = "선택한 시트에는 데이터가 없습니다.";
        private const string MSG_SELECT_SHEET = "엑셀 시트를 선택해주세요.";
        private const string MSG_NO_EXPORT_SETTINGS = "해당하는 출력 설정이 없습니다. 시트를 선택해주세요.";
        private const string MSG_NO_FILE = "파일이 존재하지 않습니다.";

        private const string STATUS_STRING_FORMAT = "Funigloo Export Table Data 로딩중... [{1}] ({2}/{3}) {0}";
        private const string PROGRASS_TOK = "■■■";


        MyExcelManager excelManager;
        SettingsManager settingsManager;


        Workbook thisWorkbook;

        Dictionary<string, Excel.Application> OpenedApplicationDic;

        #region 프로그래스 영역

        System.Action<PROGRASS_TYPE> OnPrograssInit;
        System.Action onRefreshLog;
        System.Action onFinishInit;
        System.Action<int> onPrograssUpdate;
        System.Action<string> onPrograssUpdateLog;
        public Action<int> OnPrograssUpdate { get => onPrograssUpdate; set => onPrograssUpdate = value; }
        public Action<string> OnPrograssUpdateLog { get => onPrograssUpdateLog; set => onPrograssUpdateLog = value; }
        public System.Action OnFinishInit { get => onFinishInit; set => onFinishInit = value; }
        public System.Action OnRefreshLog { get => onRefreshLog; set => onRefreshLog = value; }
        public int GetPrograssMaxValue() { return (int)PROGRASS_TYPE.finish; }

        private int InitPrograssMax = typeof(PROGRASS_TYPE_FLAG).GetEnumValues().Length - 1;
        private int InitPrograss = 0;
        private PROGRASS_TYPE_FLAG prograss = 0;
        private bool isInit = false;
        public bool IsInit { get => isInit; }
        #endregion

        #region 테스트 영역

        System.Action forceShowRibbon;
        public System.Action ForceShowRibbon { get => forceShowRibbon; set => forceShowRibbon = value; }
        #endregion

        [Flags]
        public enum PROGRASS_TYPE_FLAG
        {
            excelManager = 1 << 0,
            settingsManager = 1 << 1,
            checkRegistry = 1 << 2,
            checkSettingsPath = 1 << 3,
            finish = 1 << 10,
        }

        public enum PROGRASS_TYPE
        {
            excelManager = 0,
            settingsManager = 1,
            checkRegistry = 2,
            checkSettingsPath = 3,
            checkSettingsFileExist = 4,
            OpenSettingsFile = 5,
            InitializeProgramSettings = 6,
            InitializeProgramSetting = 7,
            InitializeExportSettings = 8,
            closeSettingsFile = 9,
            finish = 10,
        }



        string StatusBarText()
        {
            return string.Format(STATUS_STRING_FORMAT, ((PROGRASS_TYPE_FLAG)(1 << InitPrograss)).ToString(), GetPrograssBar(), InitPrograss, typeof(PROGRASS_TYPE_FLAG).GetEnumValues().Length - 2);
        }

        string GetPrograssBar()
        {
            int length = typeof(PROGRASS_TYPE_FLAG).GetEnumValues().Length - 2;
            string result = string.Empty;
            for (int i = 0; i < length; i++)
            {
                if (prograss.HasFlag((PROGRASS_TYPE_FLAG)(1 << i)))
                    result += PROGRASS_TOK;
                else
                    result += "□□□";
            }

            return result;
        }

        void UpdatePrograss(PROGRASS_TYPE type)
        {
            InitPrograss = (int)type;
            prograss |= (PROGRASS_TYPE_FLAG)(1 << InitPrograss);

            onPrograssUpdate?.Invoke(InitPrograss);

            //onPrograssTime?.Invoke(prevTime);
            OnPrograssUpdateLog?.Invoke("[Settings] " + ((PROGRASS_TYPE_FLAG)(1 << (InitPrograss))).ToString());
        }


        ExportDataProcess()
        {

        }

        public void Initialize()
        {
            try
            {
                isInit = false;

                OnPrograssInit += UpdatePrograss;
                SettingsManager.Instance.OnPrograssInit += UpdatePrograss;

                OnPrograssInit?.Invoke(PROGRASS_TYPE.excelManager);
                excelManager = MyExcelManager.Instance;

                OnPrograssInit?.Invoke(PROGRASS_TYPE.settingsManager);
                settingsManager = SettingsManager.Instance;

                OnPrograssInit?.Invoke(PROGRASS_TYPE.checkRegistry);
                RegistryKey reg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Funigloo.ExcelAddIn", false);

                if (reg != null)
                {
                    OnPrograssInit?.Invoke(PROGRASS_TYPE.checkSettingsPath);
                    settingsManager.Initialize(reg.GetValue("SettingPath").ToString());
                }
                else
                {
                    MessageBox.Show(null, MSG_NO_SETTINGS_FILE, MSGTITLE);
                }

                OnPrograssInit?.Invoke(PROGRASS_TYPE.finish);

                OnPrograssInit -= UpdatePrograss;
                SettingsManager.Instance.OnPrograssInit -= UpdatePrograss;
                onFinishInit?.Invoke();
                isInit = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public string LatestFilePath
        {
            get
            {
                return excelManager.LatestExportFilePath;
            }
        }


        public void ExportData(string sheetName, int serverType = 0, int legionType = 0)
        {
            if (thisWorkbook == null)
                thisWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            Worksheet selectedSheet = null;
            foreach (Worksheet item in thisWorkbook.Worksheets)
            {
                if (item.Name == sheetName)
                {
                    selectedSheet = item;
                    break;
                }
            }

            ExportData(selectedSheet, serverType, legionType);
        }

        public void ExportData(Worksheet selectedSheet, int serverType = 0, int legionType = 0)
        {
            thisWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (selectedSheet == null)
            {
                MessageBox.Show(MSG_SELECT_SHEET);
                return;
            }
            else if (!settingsManager.IsExportSettingExist(serverType, legionType, selectedSheet.Name))
            {
                MessageBox.Show(MSG_NO_EXPORT_SETTINGS);
                return;
            }
            else
            {

            }
            {
                //알람 끄기
                Globals.ThisAddIn.Application.DisplayAlerts = false;

                excelManager.ExportData(thisWorkbook, settingsManager.GetExportSetting(serverType, legionType, selectedSheet.Name));
                OnRefreshLog?.Invoke();
                //알람 켜기
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                return;
            }
        }

        public void ExportData(Workbook selectedWorkbook, int serverType = 0, int legionType = 0)
        {
            thisWorkbook = selectedWorkbook;

            if (selectedWorkbook.ActiveSheet == null)
            {
                MessageBox.Show(MSG_SELECT_SHEET);
                return;
            }
            else if (!settingsManager.IsExportSettingExist(serverType, legionType, selectedWorkbook.ActiveSheet.Name))
            {
                MessageBox.Show(MSG_NO_EXPORT_SETTINGS);
                return;
            }
            else
            {
                //알람 끄기
                Globals.ThisAddIn.Application.DisplayAlerts = false;

                excelManager.ExportData(selectedWorkbook, settingsManager.GetExportSetting(serverType, legionType, selectedWorkbook.ActiveSheet.Name));
                OnRefreshLog?.Invoke();
                //알람 켜기
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                return;
            }
        }

        public void ExporActivatedSheetData(int serverType = 0, int legionType = 0)
        {
            try
            {
                Workbook selectedWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
                ExportData(selectedWorkBook, serverType, legionType);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnOpenSettings(object sender, RibbonControlEventArgs e)
        {
            settingsManager.OpenSettings();
        }

        private void btnSetPath_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();

                dialog.InitialDirectory = "c:\\";
                dialog.IsFolderPicker = true;
                dialog.RestoreDirectory = true;
                dialog.Multiselect = false;

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    //Get the path of specified file
                    var filePath = dialog.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void OpenFile(string path)
        {
            if (OpenedApplicationDic == null)
                OpenedApplicationDic = new Dictionary<string, Excel.Application>();

            if (path == null)
            {
                MessageBox.Show("잘못된 경로입니다");
                return;
            }

            if (OpenedApplicationDic.ContainsKey(path))
            {
                OpenedApplicationDic.Remove(path);
            }

            if (!File.Exists(path))
            {
                MessageBox.Show("파일이 존재하지 않습니다.");
                return;
            }

            Excel.Application app = new Excel.Application();
            OpenedApplicationDic.Add(path, app);
            app.Workbooks.Open(path);
            app.Visible = true;
            app.WorkbookBeforeClose += App_WorkbookBeforeClose;
        }

        private void App_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            var tempapp = OpenedApplicationDic[Wb.FullName];
            MessageBox.Show(Wb.FullName + "이 닫혔습니다.");
            tempapp.Workbooks.Close();
            tempapp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(tempapp);
            OpenedApplicationDic.Remove(Wb.FullName);
        }

        public void btnExportData_Click()
        {
            try
            {
                Worksheet selectedSheet = Globals.ThisAddIn.Application.ActiveSheet;
                ExportData(selectedSheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void OpenSettings(object sender = null, RibbonControlEventArgs e = null)
        {
            OpenSettings();
        }

        public void OpenSettings()
        {
            settingsManager.OpenSettings();
        }

        public void SetDropDownItems(RibbonDropDown serverDropDown, RibbonDropDown legionDropDown)
        {
            var serverList = settingsManager.ProgramSetting.GetServerTypes();
            serverDropDown.Items.Clear();
            foreach (var serverType in serverList)
            {
                var newItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                newItem.Label = serverType.Value;
                serverDropDown.Items.Add(newItem);
            }

            var legionList = settingsManager.ProgramSetting.GetLegionTypes();
            legionDropDown.Items.Clear();
            foreach (var legionType in legionList)
            {
                var newItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                newItem.Label = legionType.Value;
                legionDropDown.Items.Add(newItem);
            }
        }

        public void BtnOpenFolderClick()
        {
            if (excelManager.LatestExportDirectoryPath == "")
            {
                MessageBox.Show(MSG_NO_FILE);
                return;
            }

            System.Diagnostics.Process.Start(excelManager.LatestExportDirectoryPath);
        }

        public void SetLogGalleryItems()
        {
            //var loglist = excelManager.ExportLogs;
            //int index = 0;
            //
            //if (loglist.Count == 0)
            //{
            //    logGallery.Buttons[0].Label = "Empty";
            //    logGallery.Buttons[index].Enabled = true;
            //    return;
            //}
            ////만약 로그 갤러리 버튼수보다 로그가 많아지면 로그 앞쪽 삭제
            //while (loglist.Count != logGallery.Buttons.Count)
            //{
            //    if (loglist.Count <= 0)
            //        break;
            //
            //    loglist.RemoveAt(0);
            //}
            //
            //for (; index < logGallery.Buttons.Count; ++index)
            //{
            //    if (index >= loglist.Count)
            //    {
            //        logGallery.Buttons[index].Label = "";
            //        logGallery.Buttons[index].Enabled = false;
            //        continue;
            //    }
            //    logGallery.Buttons[index].Label = loglist[index];
            //    logGallery.Buttons[index].Enabled = true;
            //}


            //
            //for (; index < logList.Count; ++index)
            //{
            //    var newItem = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            //    newItem.Click += OnLogButtonClicked;
            //    newItem.Label = logList[index];
            //    logGallery.Buttons.Add(newItem);
            //}

        }

        private void OnLogButtonClicked(object sender, RibbonControlEventArgs e)
        {
            if (sender is RibbonButton)
            {
                var button = sender as RibbonButton;
                if (button.Label == null)
                {
                    return;
                }
                else if (button.Label == "")
                {
                    return;
                }
                else if (IsErrorMessage(button.Label))
                {
                    MessageBox.Show(button.Label);
                    return;
                }
                else
                {
                    System.Diagnostics.Process.Start(button.Label);
                }
            }
        }

        private bool IsErrorMessage(string msg)
        {
            switch (msg)
            {
                case "해당하는 출력 설정이 없습니다. 시트를 확인해주세요.":
                case "활성화된 시트가 없습니다.":
                case "CombineRange 값에 포함된 시트를 찾을 수 없습니다.":
                case "CombineRange 값에 포함된 시트가 없습니다.":
                case "복사과정에서 문제가 발생했습니다.":
                case MSG_NO_DATA:
                case MSG_NO_EXPORT_SETTINGS:
                case MSG_NO_EXPORT_TARGET:
                case MSG_NO_FILE:
                case MSG_NO_SETTINGS_FILE:
                case MSG_NO_SHEET:
                case MSG_SELECT_SHEET:
                    return true;
                default:
                    return false;
            }
        }

        public void ForceShowRibbin_Toggle()
        {
            forceShowRibbon?.Invoke();
        }

        public void EditLog(System.Windows.Forms.ListBox listBox)
        {
            var loglist = excelManager.ExportLogs;

            if (loglist.Count == 0)
            {
                listBox.Items[0] = "Empty";
                return;
            }

            for (int i = 0; i < loglist.Count; ++i)
            {
                if (i >= listBox.Items.Count)
                {
                    listBox.Items.Add(loglist[i]);
                }
                else
                {
                    listBox.Items[i] = loglist[i];
                }
            }

            while (listBox.Items.Count > loglist.Count)
            {
                listBox.Items.RemoveAt(loglist.Count);
            }


            //
            //for (; index < logList.Count; ++index)
            //{
            //    var newItem = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            //    newItem.Click += OnLogButtonClicked;
            //    newItem.Label = logList[index];
            //    logGallery.Buttons.Add(newItem);
            //}
        }
    }
}
