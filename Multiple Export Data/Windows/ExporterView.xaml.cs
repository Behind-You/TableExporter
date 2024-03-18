using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Multiple_Export_Data.Windows
{
    /// <summary>
    /// Page1.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ExporterView : Page
    {
        public interface IPresenter
        {
            void Initialize();
            void InitFilterCombobox(ComboBox server, ComboBox legion);
            DataTable GetExportData_DataTable();
            void ExportAll_Async(ItemCollection items);
            DataTable GetExportData_DataTable(int selectedIndex1, int selectedIndex2);
        }
        private IPresenter presenter;

        private static Settings settings => Settings.Instance;

        public static System.Action<string> OnAddLog;
        private static CheckBox _CheckBox_SelectAll;
        public ExporterView()
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

        private void OnClicked_Export(object sender, RoutedEventArgs e)
        {
            AddLog("ExportAll_Async Start");
            presenter.ExportAll_Async(MainDataGrid.Items);
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space || e.Key == Key.Enter)
            {
                //변환 안됨
                var temp = MainDataGrid.SelectedItem as DataRowView;
                if (temp is System.Data.DataRowView)
                {
                    temp.Row["IsSelected"] = !((bool)temp.Row["IsSelected"]);
                }
            }

            if (e.Key == Key.Escape)
            {
                MainDataGrid.UnselectAllCells();
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as CheckBox;

            if(_CheckBox_SelectAll == null)
            {
                _CheckBox_SelectAll = checkBox;
            }

            var data = MainDataGrid.Items;
            foreach (var item in data)
            {
                if (item is System.Data.DataRowView)
                {
                    var temp = item as System.Data.DataRowView;
                    temp.Row["IsSelected"] = _CheckBox_SelectAll.IsChecked;
                }
            }
        }

        private void Btn_Remove_Filter_Click(object sender, RoutedEventArgs e)
        {
            ComboBox_Filter_Server.SelectedIndex = -1;
            ComboBox_Filter_Legion.SelectedIndex = -1;

            var datatable = settings.GetExportSettings_ToDataTable();
            var dataview = datatable.DefaultView;
            dataview.RowFilter = "";

            if (_CheckBox_SelectAll != null)
                _CheckBox_SelectAll.IsChecked = false;

            MainDataGrid.ItemsSource = dataview;
        }

        private void Btn_Refresh_Filter_Click(object sender, RoutedEventArgs e)
        {
            if(_CheckBox_SelectAll != null)
                _CheckBox_SelectAll.IsChecked = false;

            MainDataGrid.ItemsSource = presenter.GetExportData_DataTable(ComboBox_Filter_Server.SelectedIndex, ComboBox_Filter_Legion.SelectedIndex).DefaultView;
        }

        private void OnSelectedData(object sender, RoutedEventArgs e)
        {
            MainDataGrid.SelectedIndex = -1;
        }
    }
}
