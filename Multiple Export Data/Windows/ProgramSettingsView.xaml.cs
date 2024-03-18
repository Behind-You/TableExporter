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

namespace Multiple_Export_Data.Windows
{
    /// <summary>
    /// ProgramSettingsView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ProgramSettingsView : Page
    {
        public interface IPresenter
        {
            void Initialize();
            DataTable GetProgramSettings_DataTable();
            DataTable GetServerTypes_DataTable();
            DataTable GetLegionTypes_DataTable();
            DataTable GetTotalSheets_DataTable();
        }
        private IPresenter presenter;

        private static Settings settings => Settings.Instance;

        public static System.Action<string> OnAddLog;

        public ProgramSettingsView()
        {
            presenter = new ExporterViewPresenter();
            InitializeComponent();
            presenter.Initialize();

            AddEvent();
            Refresh();

        }

        void Refresh()
        {
            DataGrid_ProgramSettings.ItemsSource = presenter.GetProgramSettings_DataTable().DefaultView;
            DataGrid_ServerTypes.ItemsSource = presenter.GetServerTypes_DataTable().DefaultView;
            DataGrid_LegionTypes.ItemsSource = presenter.GetLegionTypes_DataTable().DefaultView;
            DataGrid_TotalSheets.ItemsSource = presenter.GetTotalSheets_DataTable().DefaultView;
        }

        void AddEvent()
        {
            App.OnSelected += OnSelectedData;
        }

        void AddLog(string message)
        {
            OnAddLog?.Invoke(message);
        }


        private void OnSelectedData(object sender, RoutedEventArgs e)
        {
            DataGrid_ProgramSettings.SelectedIndex = -1;
            DataGrid_LegionTypes.SelectedIndex = -1;
            DataGrid_ServerTypes.SelectedIndex = -1;
            DataGrid_TotalSheets.SelectedIndex = -1;
        }

        private void Grid_LegionTypes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DataGrid_LegionTypes.UnselectAllCells();
            }
        }

        private void Grid_ServerTypes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DataGrid_ServerTypes.UnselectAllCells();
            }

        }

        private void Grid_ProgramSettings_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DataGrid_ProgramSettings.UnselectAllCells();
            }
        }

        private void Grid_TotalSheets_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DataGrid_TotalSheets.UnselectAllCells();
            }
        }

        private void ListViewScrollViewer_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            //MessageBox.Show(string.Format("{0} {1} {2} {3} {4} ", sv.ActualHeight, sv.VerticalOffset, sv.ScrollableHeight, sv.ContentVerticalOffset, this.Height, this.ActualHeight));
            
        }

        private void Page_LayoutUpdated(object sender, EventArgs e)
        {
            sv.Height = this.ActualHeight - 50 - this.Margin.Top - 1;
        }
    }
}
