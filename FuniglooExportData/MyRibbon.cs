using System;
using System.Collections.Generic;
using System.Windows.Forms;
using FuniglooExportData.Properties;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuniglooExportData
{
    public partial class MyRibbon
    {
        public interface IPresenter
        {
            void Initialize();
            void btnExportData_Click();
            void OpenSettings(object sender = null, RibbonControlEventArgs e = null);
            void SetDropDownItems(RibbonDropDown serverDropDown, RibbonDropDown legionDropDown);
            void SetLogGalleryItems();
            void BtnOpenFolderClick();
            string LatestFilePath { get; }
            void ExportData(string sheetName, int serverType = 0, int legionType = 0);
            void ExportData(Excel.Worksheet sheet, int serverType = 0, int legionType = 0);
            void ExporActivatedSheetData(int serverType, int legionType);
            System.Action OnRefreshLog { get; set; }
            System.Action OnFinishInit { get; set; }
            System.Action ForceShowRibbon { get; set; }

        }

        private Dictionary<string, Workbook> Dic_OpenedWorkbooks;
        private IPresenter presenter;

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Group_ExportData.Visible = false;
            btnExportData.Visible = false;
        }

        private void OnFinishInit()
        {
            if (SettingsManager.Instance.IsInitialized)
            {
                presenter.SetDropDownItems(Dropdown_ServerType, Dropdown_LegionType);
                Group_ExportData.Visible = true;
                btnExportData.Visible = true;
                ExportDataInit.Visible = false;
            }
            else
            {
                Group_ExportData.Visible = false;
                btnExportData.Visible = false;
                ExportDataInit.Visible = true;
            }
        }

        private void OnWorkbookClose(Workbook Wb, ref bool Cancel)
        {
            if (Dic_OpenedWorkbooks.ContainsKey(Wb.Name))
            {
                Dic_OpenedWorkbooks.Remove(Wb.Name);
            }

            if (Dic_OpenedWorkbooks.Count == 0)
            {
                RemoveEvents();
            }
        }

        private void Init()
        {
            if (!ExportDataProcess.Instance.IsInit)
                ExportDataProcess.Instance.Initialize();
            presenter = ExportDataProcess.Instance;
            presenter.OnFinishInit += OnFinishInit;

            //파일 오픈 후 로딩
            OnWorkBookOpen(Globals.ThisAddIn.Application.ActiveWorkbook);
            Globals.ThisAddIn.Application.WorkbookBeforeClose += OnWorkbookClose;
        }

        private void OnWorkBookOpen(Workbook Wb)
        {
            try
            {
                if (Wb.Name == "Settings.xlsx")
                {
                    Group_ExportData.Visible = false;
                    return;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            if (Dic_OpenedWorkbooks == null)
                Dic_OpenedWorkbooks = new Dictionary<string, Workbook>();

            if (Dic_OpenedWorkbooks.Count == 0)
            {
                AddEvents();
            }

            if (Dic_OpenedWorkbooks.ContainsKey(Wb.Name))
            {
                Dic_OpenedWorkbooks.Add(Wb.Name, Wb);
            }
            OnFinishInit();
        }

        void AddEvents()
        {
            Dropdown_ServerType.Buttons[0].Click += presenter.OpenSettings;
            Dropdown_LegionType.Buttons[0].Click += presenter.OpenSettings;
            presenter.ForceShowRibbon += ToggleForceShowRibbon;

        }

        void RemoveEvents()
        {
            Dropdown_ServerType.Buttons[0].Click -= presenter.OpenSettings;
            Dropdown_LegionType.Buttons[0].Click -= presenter.OpenSettings;
        }


        private void btnExportData_Click(object sender, RibbonControlEventArgs e)
        {
            if(presenter == null)
            {
                Init();
            }

            presenter.ExporActivatedSheetData(Dropdown_ServerType.SelectedItemIndex, Dropdown_LegionType.SelectedItemIndex);
        }

        private void btnOpenSettings(object sender, RibbonControlEventArgs e)
        {
            if (presenter == null)
            {
                Init();
            }
            presenter.OpenSettings();
        }

        private void ToggleForceShowRibbon()
        {
            if (presenter == null)
            {
                Init();
            }

            Group_ExportData.Visible = !Group_ExportData.Visible;
        }

        private void btnInit(object sender, RibbonControlEventArgs e)
        {
            if (presenter == null)
            {
                Init();
            }
        }
    }
}
