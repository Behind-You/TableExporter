using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Windows.Controls;
using Multiple_Export_Data.Windows;

namespace Multiple_Export_Data
{
    public class MainWindowPresenter : MainWindow.IPresenter
    {
        private Settings settings;

        private System.Action<string> _OnAddLog;
        private ExporterView _ExporterView;
        private ExportSettingsView _ExportSettingsView;
        private ProgramSettingsView _ProgramSettingsView;
        private TotalSheetView _TotalSheetView;

        public Action<string> OnAddLog { get => _OnAddLog; set => _OnAddLog += value; }
        public ExporterView ExporterView
        {
            get
            {
                if (_ExporterView == null)
                    _ExporterView = new ExporterView();
                return _ExporterView;
            }
        }

        public ExportSettingsView ExportSettingsView
        {
            get
            {
                if (_ExportSettingsView == null)
                    _ExportSettingsView = new ExportSettingsView();
                return _ExportSettingsView;
            }
        }

        public ProgramSettingsView ProgramSettingsView
        {
            get
            {
                if (_ProgramSettingsView == null)
                    _ProgramSettingsView = new ProgramSettingsView();
                return _ProgramSettingsView;
            }
        }

        public TotalSheetView TotalSheetView
        {
            get
            {
                if (_TotalSheetView == null)
                    _TotalSheetView = new TotalSheetView();
                return _TotalSheetView;
            }
        }

        public void Initialize()
        {
            settings = Settings.Instance;
            settings.OnChangeStatus += ShowSettingsStatusLog;
            settings.OnMessage += ShowMessage;

            TableExporter.Instance.OnAddLog_TableExport += ShowTableExporterMessage;

            ExcelManager.OnMessage += ShowExcelMessage;
            ExcelManager.OnExportFinish += ShowExcelMessage;
            ExcelManager.OnLog += ShowExcelMessage;            
        }

        public void OpenSettings()
        {
            settings.OpenSettings();
        }

        public void ShowSettingsStatusLog(Settings.Current_Status status)
        {
            string message = AddTimeStamp(status.ToString());
            message = "[Settings] " + message;
            Console.WriteLine(message);
            _OnAddLog?.Invoke(message);
        }

        public void ShowMessage(string message)
        {
            Console.WriteLine(message);
            _OnAddLog?.Invoke(AddTimeStamp(message));
        }

        public void ShowExcelMessage(string message)
        {
            message = AddTimeStamp(message);
            message = "[Excel] " + message;
            Console.WriteLine(message);
            _OnAddLog?.Invoke(message);
        }

        public void ShowTableExporterMessage(Queue<string> messages)
        {
            while (messages.Count != 0)
            {
                string message = messages.Dequeue();
                message = "[TableExporter] " + message;
                Console.WriteLine(message);
                _OnAddLog?.Invoke(message);

            }
        }

        public string AddTimeStamp(string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            const string FORMAT = "[{0}] {1}";
            return string.Format(FORMAT, timestamp, message);
        }
    }

    public class ExporterViewPresenter : ExporterView.IPresenter, ExportSettingsView.IPresenter, ProgramSettingsView.IPresenter, TotalSheetView.IPresenter
    {
        private Settings settings;
        private TableExporter tableExporter;

        private System.Action<string> _OnAddLog;

        public void Initialize()
        {
            settings = Settings.Instance;
            tableExporter = TableExporter.Instance;
        }

        public void ExportAll(ItemCollection items)
        {
            tableExporter.ExportAll(items);
        }

        public void ExportAll_Async(ItemCollection items)
        {
            tableExporter.ExportAll_Async(items);
        }
        public void InitFilterCombobox(ComboBox Filter_Server, ComboBox Filter_Legion)
        {
            settings.GetProgramSetting().GetServerTypes().ForEach(x => Filter_Server.Items.Add(x.Value));
            settings.GetProgramSetting().GetLegionTypes().ForEach(x => Filter_Legion.Items.Add(x.Value));
        }

        public System.Data.DataTable GetExportData_DataTable()
        {
            return settings.GetExportSettings_ToDataTable();
        }

        public System.Data.DataTable GetExportData_DataTable(int serverType, int legionType)
        {
            var datatable = settings.GetExportSettings_ToDataTable();
            var dataview = datatable.DefaultView;
            try
            {
                if(serverType != -1 && legionType != -1)
                    dataview.RowFilter = $"SERVER_TYPE = '{settings.GetServerTypeValue(serverType)}' AND LEGION_TYPE = '{settings.GetLegionTypeValue(legionType)}'";
            }
            catch (Exception ex)
            {
                _OnAddLog?.Invoke(ex.Message);
                dataview.RowFilter = "";
            }
            return datatable;
        }

        public DataTable GetProgramSettings_DataTable()
        {
            return settings.GetProgramSettings_ProgramSettings_ToDataTable();
        }

        public DataTable GetServerTypes_DataTable()
        {
            return settings.GetProgramSettings_ServerTypes_ToDataTable();
        }

        public DataTable GetLegionTypes_DataTable()
        {
            return settings.GetProgramSettings_LegionTypes_ToDataTable();
        }

        public DataTable GetTotalSheets_DataTable()
        {
            return settings.GetProgramSettings_TotalSheet_ToDataTable();
        }

        public List<ProgramSetting.TotalSheetSettingsInfo> GetTotalSheetSettings()
        {
            return settings.GetProgramSetting().GetTotalSheetSettings();
        }

        public void OpenWorkbook(string path)
        {
            settings.OpenWorkBook(path);
        }
    }
}
