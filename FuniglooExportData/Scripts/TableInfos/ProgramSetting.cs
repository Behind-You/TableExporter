using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace FuniglooExportData
{
    /// <summary>
    /// 프로그램 설정 정보
    /// </summary>
    public class ProgramSetting
    {
        public ProgramSettingInfo ExportSettingRange => ProgramSettings[0];
        public ProgramSettingInfo PathFormat => ProgramSettings[1];
        public ProgramSettingInfo DefaultExtension => ProgramSettings[2];

        public struct ProgramSettingInfo
        {
            public int ID;
            public string Name;
            public CellValue Value;
        }

        public struct ServerTypeInfo
        {
            public int ID;
            public string Name;
            public string Value;
        }

        public struct LegionTypeInfo
        {
            public int ID;
            public string Name;
            public string Value;
        }

        private readonly List<ProgramSettingInfo> ProgramSettings;
        private readonly List<ServerTypeInfo> ServerTypes;
        private readonly List<LegionTypeInfo> LegionTypes;

        private readonly Dictionary<int, ProgramSettingInfo> ProgramSettingsDict;
        private readonly Dictionary<int, ServerTypeInfo> ServerTypesDict;
        private readonly Dictionary<int, LegionTypeInfo> LegionTypesDict;

        public ProgramSetting(Excel.Worksheet _settingSheet)
        {
            ProgramSettings = new List<ProgramSettingInfo>();
            ServerTypes = new List<ServerTypeInfo>();
            LegionTypes = new List<LegionTypeInfo>();

            ProgramSettingsDict = new Dictionary<int, ProgramSettingInfo>();
            ServerTypesDict = new Dictionary<int, ServerTypeInfo>();
            LegionTypesDict = new Dictionary<int, LegionTypeInfo>();

            ResetData();
            Initialize(_settingSheet);
        }

        public void ResetData()
        {
            ProgramSettings.Clear();
            ServerTypes.Clear();
            LegionTypes.Clear();

            ProgramSettingsDict.Clear();
            ServerTypesDict.Clear();
            LegionTypesDict.Clear();
        }

        public void Initialize(Excel.Worksheet _settingSheet)
        {
            SetProgramSettings(_settingSheet.Range["A:D"]);
            SetServerTypes(_settingSheet.Range["E:G"]);
            SetLegionTypes(_settingSheet.Range["H:J"]);
        }

        public ProgramSettingInfo GetProgramSetting(int id)
        {
            if (ProgramSettingsDict.ContainsKey(id))
                return ProgramSettingsDict[id];
            else
                return new ProgramSettingInfo();
        }

        public ServerTypeInfo GetServerType(int id)
        {
            if (ServerTypesDict.ContainsKey(id))
                return ServerTypesDict[id];
            else
                return new ServerTypeInfo();
        }

        public LegionTypeInfo GetLegionType(int id)
        {
            if (LegionTypesDict.ContainsKey(id))
                return LegionTypesDict[id];
            else
                return new LegionTypeInfo();
        }

        public List<ProgramSettingInfo> GetProgramSettings()
        {
            return ProgramSettings;
        }

        public List<ServerTypeInfo> GetServerTypes()
        {
            return ServerTypes;
        }

        public List<LegionTypeInfo> GetLegionTypes()
        {
            return LegionTypes;
        }


        void SetProgramSettings(Range data)
        {
            for (int row = 1; row < data.Rows.Count; row++)
            {
                Range rows = data.Rows[row];
                List<CellData> cells = CellData.RangeToCellDataList(rows, true);

                if (cells.Count == 0)
                    break;
                else if (cells[0] == null)
                    break;
                else if (cells[0].Value == null)
                    break;
                else if (cells[0].Type != null && cells[0].Type == typeof(string) && cells[0].Value.ToString == "ProgramSettingsID")
                    continue;
                ProgramSettingInfo setting = new ProgramSettingInfo();
                setting.ID = cells[0].Value.ToInt;
                setting.Name = cells[1].Value.ToString;
                setting.Value = new CellValue();
                setting.Value.Type = cells[2].Value.ToType;
                setting.Value.Value = cells[3].Value.ToString;

                ProgramSettings.Add(setting);
                ProgramSettingsDict.Add(setting.ID, setting);
            }
        }

        void SetServerTypes(Range data)
        {
            for (int row = 1; row < data.Rows.Count; row++)
            {
                Range rows = data.Rows[row];
                List<CellData> cells = CellData.RangeToCellDataList(rows, true);

                if (cells.Count == 0)
                    break;
                else if (cells[0] == null)
                    break;
                else if (cells[0].Value == null)
                    break;
                else if (cells[0].Type != null && cells[0].Type == typeof(string) && cells[0].Value.ToString == "ServerTypesName")
                    continue;
                ServerTypeInfo setting = new ServerTypeInfo();
                setting.Name = cells[0].Value.ToString;
                setting.Value = cells[1].Value.ToString;
                setting.ID = cells[2].Value.ToInt;

                ServerTypes.Add(setting);
                ServerTypesDict.Add(setting.ID, setting);
            }
        }

        void SetLegionTypes(Range data)
        {
            for (int row = 1; row < data.Rows.Count; row++)
            {
                Range rows = data.Rows[row];
                List<CellData> cells = CellData.RangeToCellDataList(rows, true);

                if (cells.Count == 0)
                    break;
                else if (cells[0] == null)
                    break;
                else if (cells[0].Value == null)
                    break;
                else if (cells[0].Type != null && cells[0].Type == typeof(string) && cells[0].Value.ToString == "LegionTypesName")
                    continue;
                LegionTypeInfo setting = new LegionTypeInfo();
                setting.Name = cells[0].Value.ToString;
                setting.Value = cells[1].Value.ToString;
                setting.ID = cells[2].Value.ToInt;

                LegionTypes.Add(setting);
                LegionTypesDict.Add(setting.ID, setting);
            }
        }
    }
}
