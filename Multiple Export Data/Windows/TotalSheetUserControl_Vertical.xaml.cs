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

namespace Multiple_Export_Data.Windows
{
    /// <summary>
    /// TotalSheetUserControl_Vertical.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class TotalSheetUserControl_Vertical : UserControl
    {
        public string FullName;
        public string TitleName;

        public TotalSheetUserControl_Vertical()
        {
            InitializeComponent();
            Text_Workbook_Name.Text = TitleName;
        }

        void ButtonClick(object sender, RoutedEventArgs e) => RaiseEvent(new RoutedEventArgs(ElementClickEvent, sender));

        public string Title
        {
            get { return (string)GetValue(TitleProperty); }
            set { SetValue(TitleProperty, value); }
        }

        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register("Title", typeof(string), typeof(TotalSheetUserControl), new PropertyMetadata(""));

        public event RoutedEventHandler ElementClick
        {
            add { AddHandler(ElementClickEvent, value); }
            remove { RemoveHandler(ElementClickEvent, value); }
        }

        public static readonly RoutedEvent ElementClickEvent =
            EventManager.RegisterRoutedEvent("ElementClick", RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(TotalSheetUserControl));

    }
}
