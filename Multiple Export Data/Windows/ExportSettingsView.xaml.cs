using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Multiple_Export_Data.Windows
{
    /// <summary>
    /// ExportSettingsView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ExportSettingsView : Page
    {
        public interface IPresenter
        {
            void Initialize();
            void InitFilterCombobox(ComboBox server, ComboBox legion);
            DataTable GetExportData_DataTable();
            DataTable GetExportData_DataTable(int selectedIndex1, int selectedIndex2);
        }
        private IPresenter presenter;

        private static Settings settings => Settings.Instance;

        public static System.Action<string> OnAddLog;

        public ExportSettingsView()
        {
            presenter = new ExporterViewPresenter();
            InitializeComponent();
            presenter.Initialize();

            AddEvent();
            Refresh();
        }

        void Refresh()
        {
            ComboBox_Filter_Server.Items.Clear();
            ComboBox_Filter_Legion.Items.Clear();
            presenter.InitFilterCombobox(ComboBox_Filter_Server, ComboBox_Filter_Legion);

            MainDataGrid.ItemsSource = presenter.GetExportData_DataTable().DefaultView;
        }

        void AddEvent()
        {
            App.OnSelected += OnSelectedData;
        }

        void AddLog(string message)
        {
            OnAddLog?.Invoke(message);
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                MainDataGrid.UnselectAllCells();
            }
        }

        private void Btn_Remove_Filter_Click(object sender, RoutedEventArgs e)
        {
            ComboBox_Filter_Server.SelectedIndex = -1;
            ComboBox_Filter_Legion.SelectedIndex = -1;

            var datatable = settings.GetExportSettings_ToDataTable();
            var dataview = datatable.DefaultView;
            dataview.RowFilter = "";

            MainDataGrid.ItemsSource = dataview;
        }

        private void Btn_Refresh_Filter_Click(object sender, RoutedEventArgs e)
        {
            MainDataGrid.ItemsSource = presenter.GetExportData_DataTable(ComboBox_Filter_Server.SelectedIndex, ComboBox_Filter_Legion.SelectedIndex).DefaultView;
        }

        private void OnSelectedData(object sender, RoutedEventArgs e)
        {
            MainDataGrid.SelectedIndex = -1;
        }
    }
}
