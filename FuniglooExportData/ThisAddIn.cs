namespace FuniglooExportData
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _myCustomTaskPane;
        private MyUserControl _myUserControl;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Windows.Forms.Application.ThreadExit += new System.EventHandler(SettingsManager.Instance.OnShutdownSettingApplication);
            //파일 오픈 후 로딩
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            SettingsManager.Instance.OnShutdownSettingApplication();
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
