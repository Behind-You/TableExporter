using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using System.Windows.Controls.Primitives;
using Multiple_Export_Data.Windows;

namespace Multiple_Export_Data
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public interface IPresenter
        {
            ExporterView ExporterView { get; }
            ExportSettingsView ExportSettingsView { get; }
            ProgramSettingsView ProgramSettingsView { get; }
            TotalSheetView TotalSheetView { get;}
            Action<string> OnAddLog { get; set; }

            void Initialize();
            void OpenSettings();
        }

        private Style LogTextStyle;

        private static bool isInit = false;

        private static Queue<string> logQueue = new Queue<string>();
        private static readonly object m_lockObj = new object();

        private static IPresenter presenter;

        public MainWindow()
        {
            presenter = new MainWindowPresenter();
            Visibility = Visibility.Hidden;
            AddEvent();
            Startup startup = new Startup();
            startup.ShowDialog();
            Visibility = Visibility.Visible;

            InitializeComponent();
            LogTextStyle = Application.Current.FindResource("LogText") as Style;
            Init();
        }

        void AddEvent()
        {
            this.Activated += MainWindow_Activated;
            presenter.OnAddLog += AddLog;
        }

        private void MainWindow_Activated(object sender, EventArgs e)
        {
            Init();
        }

        void Init()
        {
            if (isInit == true)
                return;
            presenter.Initialize();

            DispatcherTimer timer = new DispatcherTimer();

            timer.Interval = TimeSpan.FromMilliseconds(100);
            timer.Tick += new EventHandler(timer_Tick);
            timer.Start();

            MainView.Navigate(presenter.ExporterView);

            isInit = true;
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            if (logQueue.Count > 0)
            {
                lock(m_lockObj)
                {
                    for(int i = 0; logQueue.Count > 0; i++)
                    {
                        ListBox_Log.Items.Add(new TextBlock() { Text = logQueue.Dequeue(), Style = LogTextStyle });
                        
                        //자동 스크롤
                        ListBox_Log.SelectedIndex = ListBox_Log.Items.Count - 1;
                        ListBox_Log.ScrollIntoView(ListBox_Log.SelectedItem);
                        ListBox_Log.SelectedIndex = -1;
                    }
                }
            }
        }

        void AddLog(string message)
        {
            lock (m_lockObj)
            {
                logQueue.Enqueue(message);
            }
        }

        #region 이벤트

        private void MenuBtn_OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            presenter.OpenSettings();
        }

        private void MenuBtn_ShowTotalSheets_Click(object sender, RoutedEventArgs e)
        {
            MainView.Navigate(presenter.TotalSheetView);
        }

        private void NotAddedYet(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("추가 예정");
        }

        private void ThumbDragStarted(object sender, RoutedEventArgs e)
        {        
            this.Cursor = Cursors.SizeWE;
        }

        private void thumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            double adjustX = MenuBoarder.Width + e.HorizontalChange;
            adjustX = adjustX > MenuBoarder.MinWidth ? adjustX : MenuBoarder.MinWidth;
            MenuBoarder.Width = adjustX;
        }

        private void thumbtest_DragCompleted(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void Btn_Settings_Refresh_Filter_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Btn_Settings_Remove_Filter_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ShowExporter(object sender, RoutedEventArgs e)
        {
            MainView.Navigate(presenter.ExporterView);
        }

        private void ShowExporterSettings(object sender, RoutedEventArgs e)
        {
            MainView.Navigate(presenter.ExportSettingsView);
        }

        private void ShowProgramSettings(object sender, RoutedEventArgs e)
        {
            MainView.Navigate(presenter.ProgramSettingsView);
        }

        private void OnSizeChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ListBox_Log.SelectedIndex = ListBox_Log.Items.Count - 1;
            ListBox_Log.SelectedIndex = -1;
        }

        #endregion
    }
}
