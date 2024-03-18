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
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Win32;

namespace Multiple_Export_Data
{
    /// <summary>
    /// Startup.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Startup : Window
    {
        private bool IsStartInit = false;
        private Queue<(string log, int percentage)> prograssQueue = new Queue<(string log, int percentage)>();
        public Startup()
        {
            Settings.Instance.OnChangeStatus += ShowSettingStatusLog;
            Settings.Instance.OnMessage += ShowMessage;

            InitializeComponent();

            DispatcherTimer timer = new DispatcherTimer();

            timer.Interval = TimeSpan.FromMilliseconds(10);
            timer.Tick += new EventHandler(timer_Tick);
            timer.Start();
        }

        private void Startup_Deactivated(object sender, EventArgs e)
        {
        }

        private void ShowMessage(string message)
        {
            MessageBox.Show(message);
        }

        private void ShowSettingStatusLog(Settings.Current_Status status)
        {
            string log = string.Empty;
            switch (status)
            {
                case Settings.Current_Status.NotInitialized:
                    log = "초기화 준비중입니다.";
                    break;
                case Settings.Current_Status.checkSettingsPath:
                    log = "Settings.xlsx 파일 경로 확인중입니다.";
                    break;
                case Settings.Current_Status.checkSettingsFileExist:
                    log = "Settings.xlsx 파일 존재 확인중입니다.";
                    break;
                case Settings.Current_Status.OpenSettingsFile:
                    log = "Settings.xlsx 파일 접근중입니다";
                    break;
                case Settings.Current_Status.LoadSettingsFile:
                    log = "Settings.xlsx 파일 로드중입니다";
                    break;
                case Settings.Current_Status.InitializeProgramSetting:
                    log = "Program Setting 로드중입니다";
                    break;
                case Settings.Current_Status.InitializeExportSettings:
                    log = "Export Setting 로드중입니다";
                    break;
                case Settings.Current_Status.closeSettingsFile:
                    log = "Settings.xlsx 파일 종료중입니다.";
                    break;
            }
            prograssQueue.Enqueue((log, 100 * status.GetHashCode() / Settings.Current_Status.closeSettingsFile.GetHashCode()));
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            if (prograssQueue.Count > 0)
            {
                var prograss = prograssQueue.Dequeue();
                TextBlock_Log.Text = prograss.log;
                PrograssBar_Loading.Value = prograss.percentage;

                if (prograss.percentage == 100)
                {
                    this.Close();
                }
            }

            if(IsStartInit)
            {
                IsStartInit = false;
                RegistryKey reg = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\Funigloo.ExcelAddIn", false);
                if (reg == null)
                {
                    MessageBox.Show("레지스트리 값이 없습니다.");
                    this.Close();
                    return;
                }
                Settings.Instance.Initialize(reg.GetValue("SettingPath").ToString());
            }
        }

        private void MainWindow_Activated(object sender, EventArgs e)
        {
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            if(!IsStartInit)
            {
                IsStartInit = true;
            }
        }
    }
}
