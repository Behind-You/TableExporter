using System;
using System.Collections.Generic;
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
using System.Data;

namespace Multiple_Export_Data.Windows
{
    /// <summary>
    /// TotalSheetView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class TotalSheetView : Page
    {
        public interface IPresenter
        {
            void Initialize();
            List<ProgramSetting.TotalSheetSettingsInfo> GetTotalSheetSettings();
            void OpenWorkbook(string path);
        }

        IPresenter presenter;

        public TotalSheetView()
        {
            presenter = new ExporterViewPresenter();
            presenter.Initialize();
            InitializeComponent();
            StackPanel_TotalSheets.Children.Clear();
            foreach(var item in presenter.GetTotalSheetSettings())
            {
                var child = new TotalSheetUserControl_Vertical();
                child.TitleName = item.Name;
                child.FullName = item.Path;
                child.ElementClick += Child_ElementClick;
                child.Text_Workbook_Name.Text = item.Name;
                StackPanel_TotalSheets.Children.Add(child);
            }
        }

        private void Child_ElementClick(object sender, RoutedEventArgs e)
        {
            if(Equals(sender, e.Source as TotalSheetUserControl_Vertical))
                presenter.OpenWorkbook((e.Source as TotalSheetUserControl_Vertical).FullName);
        }

        private void Page_LayoutUpdated(object sender, EventArgs e)
        {
            sv.Height = this.ActualHeight - 50 - this.Margin.Top - 1;
            sv.Width = this.ActualWidth - 50 - this.Margin.Left - 1;
        }
    }
}
