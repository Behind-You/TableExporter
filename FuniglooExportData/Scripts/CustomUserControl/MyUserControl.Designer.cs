namespace FuniglooExportData
{
    partial class MyUserControl
    {
        /// <summary> 
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary> 
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.Group_Logs = new System.Windows.Forms.GroupBox();
            this.ListBox_Log = new System.Windows.Forms.ListBox();
            this.Group_ProgramState = new System.Windows.Forms.GroupBox();
            this.PGBar_Settings = new System.Windows.Forms.ProgressBar();
            this.btnRefreshSettings = new System.Windows.Forms.Button();
            this.btnOpenSettings = new System.Windows.Forms.Button();
            this.Group_TEST = new System.Windows.Forms.GroupBox();
            this.btn_ShowMainRibbon = new System.Windows.Forms.Button();
            this.Group_Logs.SuspendLayout();
            this.Group_ProgramState.SuspendLayout();
            this.Group_TEST.SuspendLayout();
            this.SuspendLayout();
            // 
            // Group_Logs
            // 
            this.Group_Logs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Group_Logs.AutoSize = true;
            this.Group_Logs.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Group_Logs.Controls.Add(this.ListBox_Log);
            this.Group_Logs.Location = new System.Drawing.Point(0, 124);
            this.Group_Logs.Margin = new System.Windows.Forms.Padding(0);
            this.Group_Logs.MinimumSize = new System.Drawing.Size(200, 350);
            this.Group_Logs.Name = "Group_Logs";
            this.Group_Logs.Padding = new System.Windows.Forms.Padding(0);
            this.Group_Logs.Size = new System.Drawing.Size(200, 357);
            this.Group_Logs.TabIndex = 1;
            this.Group_Logs.TabStop = false;
            this.Group_Logs.Text = "Logs";
            // 
            // ListBox_Log
            // 
            this.ListBox_Log.FormattingEnabled = true;
            this.ListBox_Log.ItemHeight = 12;
            this.ListBox_Log.Location = new System.Drawing.Point(5, 15);
            this.ListBox_Log.Margin = new System.Windows.Forms.Padding(0);
            this.ListBox_Log.MinimumSize = new System.Drawing.Size(190, 304);
            this.ListBox_Log.Name = "ListBox_Log";
            this.ListBox_Log.Size = new System.Drawing.Size(190, 328);
            this.ListBox_Log.TabIndex = 0;
            this.ListBox_Log.SelectedIndexChanged += new System.EventHandler(this.ListBox_Log_SelectedIndexChanged);
            this.ListBox_Log.DoubleClick += new System.EventHandler(this.ListBox_Log_DoubleClick);
            // 
            // Group_ProgramState
            // 
            this.Group_ProgramState.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Group_ProgramState.AutoSize = true;
            this.Group_ProgramState.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Group_ProgramState.Controls.Add(this.PGBar_Settings);
            this.Group_ProgramState.Controls.Add(this.btnRefreshSettings);
            this.Group_ProgramState.Controls.Add(this.btnOpenSettings);
            this.Group_ProgramState.Location = new System.Drawing.Point(0, 0);
            this.Group_ProgramState.Margin = new System.Windows.Forms.Padding(0);
            this.Group_ProgramState.MinimumSize = new System.Drawing.Size(200, 110);
            this.Group_ProgramState.Name = "Group_ProgramState";
            this.Group_ProgramState.Padding = new System.Windows.Forms.Padding(5);
            this.Group_ProgramState.Size = new System.Drawing.Size(200, 124);
            this.Group_ProgramState.TabIndex = 0;
            this.Group_ProgramState.TabStop = false;
            this.Group_ProgramState.Text = "Addin Status";
            // 
            // PGBar_Settings
            // 
            this.PGBar_Settings.Location = new System.Drawing.Point(10, 79);
            this.PGBar_Settings.Name = "PGBar_Settings";
            this.PGBar_Settings.Size = new System.Drawing.Size(180, 23);
            this.PGBar_Settings.TabIndex = 2;
            // 
            // btnRefreshSettings
            // 
            this.btnRefreshSettings.Location = new System.Drawing.Point(10, 49);
            this.btnRefreshSettings.Name = "btnRefreshSettings";
            this.btnRefreshSettings.Size = new System.Drawing.Size(180, 23);
            this.btnRefreshSettings.TabIndex = 1;
            this.btnRefreshSettings.Text = "Refresh Settings";
            this.btnRefreshSettings.UseVisualStyleBackColor = true;
            this.btnRefreshSettings.Click += new System.EventHandler(this.btnRefreshSettings_Click);
            // 
            // btnOpenSettings
            // 
            this.btnOpenSettings.Location = new System.Drawing.Point(10, 20);
            this.btnOpenSettings.Name = "btnOpenSettings";
            this.btnOpenSettings.Size = new System.Drawing.Size(180, 23);
            this.btnOpenSettings.TabIndex = 0;
            this.btnOpenSettings.Text = "Open Settings";
            this.btnOpenSettings.UseVisualStyleBackColor = true;
            this.btnOpenSettings.Click += new System.EventHandler(this.btnOpenSettings_Click);
            // 
            // Group_TEST
            // 
            this.Group_TEST.Controls.Add(this.btn_ShowMainRibbon);
            this.Group_TEST.Location = new System.Drawing.Point(4, 485);
            this.Group_TEST.Name = "Group_TEST";
            this.Group_TEST.Size = new System.Drawing.Size(196, 108);
            this.Group_TEST.TabIndex = 2;
            this.Group_TEST.TabStop = false;
            this.Group_TEST.Text = "Test Area";
            // 
            // btn_ShowMainRibbon
            // 
            this.btn_ShowMainRibbon.Location = new System.Drawing.Point(6, 20);
            this.btn_ShowMainRibbon.Name = "btn_ShowMainRibbon";
            this.btn_ShowMainRibbon.Size = new System.Drawing.Size(180, 23);
            this.btn_ShowMainRibbon.TabIndex = 3;
            this.btn_ShowMainRibbon.Text = "Toggle Main Ribbbon";
            this.btn_ShowMainRibbon.UseVisualStyleBackColor = true;
            this.btn_ShowMainRibbon.Click += new System.EventHandler(this.btn_ShowMainRibbon_Click);
            // 
            // MyUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.Group_TEST);
            this.Controls.Add(this.Group_Logs);
            this.Controls.Add(this.Group_ProgramState);
            this.Name = "MyUserControl";
            this.Size = new System.Drawing.Size(374, 1075);
            this.Group_Logs.ResumeLayout(false);
            this.Group_ProgramState.ResumeLayout(false);
            this.Group_TEST.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox Group_Logs;
        private System.Windows.Forms.ListBox ListBox_Log;
        private System.Windows.Forms.GroupBox Group_ProgramState;
        private System.Windows.Forms.ProgressBar PGBar_Settings;
        private System.Windows.Forms.Button btnRefreshSettings;
        private System.Windows.Forms.Button btnOpenSettings;
        private System.Windows.Forms.GroupBox Group_TEST;
        private System.Windows.Forms.Button btn_ShowMainRibbon;
    }
}
