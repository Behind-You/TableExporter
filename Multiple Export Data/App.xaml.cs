using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Excel = Microsoft.Office.Interop.Excel;

namespace Multiple_Export_Data
{
    /// <summary>
    /// App.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class App : Application
    {
        public static System.Action<object, RoutedEventArgs> OnSelected;

        private void OnSelectedData(object sender, RoutedEventArgs e)
        {
            //var temp = sender as DataGridRow;
            //if (temp.Item is System.Data.DataRowView)
            //{
            //    var temp2 = temp.Item as System.Data.DataRowView;
            //    temp2.Row["IsSelected"] = !((bool)temp2.Row["IsSelected"]);
            //}
            //
            //OnSelected?.Invoke(sender, e);
        }
    }
}
