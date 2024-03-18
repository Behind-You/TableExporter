using System;
using System.Windows.Forms;

namespace FuniglooExportData
{
    public partial class MyUserControl : UserControl
    {
        public interface IPresenter
        {
            void Initialize();
            void OpenSettings();
            System.Action<int> OnPrograssUpdate { get; set; }
            System.Action<string> OnPrograssUpdateLog { get; set; }
            System.Action OnRefreshLog { get; set; }
            int GetPrograssMaxValue();
            void EditLog(ListBox listBox);

            void OpenFile(string path);

            void ForceShowRibbin_Toggle();
        }

        IPresenter presenter;

        public MyUserControl(IPresenter _presenter)
        {
            InitializeComponent();
            Initialize(_presenter);
        }

        public void Initialize(IPresenter _presenter)
        {
            if (presenter == null)
            {
                presenter = _presenter;
                presenter.OnPrograssUpdate += PrograssUpdate;
                presenter.OnPrograssUpdateLog += PrograssUpdateLog;
            }

            PGBar_Settings.Maximum = presenter.GetPrograssMaxValue();
            this.ResizeRedraw = true;
            this.Resize += MyUserControl_Resize;
        }

        private void Btn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("테스트");
        }

        private void MyUserControl_Resize(object sender, EventArgs e)
        {
            PGBar_Settings.Width = this.Width - 20;
            btnOpenSettings.Width = this.Width - 20;
            btnRefreshSettings.Width = this.Width - 20;
            ListBox_Log.Width = this.Width - 20;
            Group_Logs.Width = this.Width - 20;
        }

        private void PrograssUpdate(int value)
        {
            PGBar_Settings.Value = value;
        }

        private void PrograssUpdateLog(string log)
        {
            ListBox_Log.Items.Add(log);
        }

        private void btnOpenSettings_Click(object sender, EventArgs e)
        {
            presenter.OpenSettings();
        }

        private void btnRefreshSettings_Click(object sender, EventArgs e)
        {
            presenter.Initialize();
        }

        private void ListBox_Log_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ListBox_Log_DoubleClick(object sender, EventArgs e)
        {
            string path = ListBox_Log.SelectedItem.ToString();
            string tok = MyExcelManager.LOG_EXPORT;
            if (!path.Contains(tok))
            {
                MessageBox.Show("출력 로그가 아닙니다.");
            }
            else
            {
                presenter.OpenFile(path.Replace(tok, ""));
            }

        }

        private void btn_ShowMainRibbon_Click(object sender, EventArgs e)
        {
            presenter.ForceShowRibbin_Toggle();
        }
    }
}
