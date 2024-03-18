using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace Multiple_Export_Data
{
    public class ExportSetting
    {
        public enum ExportParameter
        {
            INDEX = 0,
            SERVER_TYPE,
            LEGION_TYPE,
            NAME,
            PATH,
            SOURCE_NAME,
            SOURCE_RANGE,
            TARGET_RANGE,
            COMBINE_COUNT,
            COMBINE_RANGES,
        }

        public readonly double Index;
        public readonly int ServerType;
        public readonly int LegionType;
        public readonly string Name;
        public readonly string Path;
        public readonly string SourceName;
        public readonly string SourceRange;
        public readonly string TargetRange;
        public readonly double CombineCount;
        public readonly List<(string SheetName, string Range, string Origin)> CombinRanges;

        public ExportSetting(List<CellData> datas)
        {
            this.Index = datas[(int)ExportParameter.INDEX].Value.ToInt;
            this.ServerType = datas[(int)ExportParameter.SERVER_TYPE].Value.ToInt;
            this.LegionType = datas[(int)ExportParameter.LEGION_TYPE].Value.ToInt;
            this.Name = datas[(int)ExportParameter.NAME].Value.ToString;
            this.Path = datas[(int)ExportParameter.PATH].Value.ToString;
            this.SourceName = datas[(int)ExportParameter.SOURCE_NAME].Value.ToString;
            this.SourceRange = datas[(int)ExportParameter.SOURCE_RANGE].Value.ToString;
            this.TargetRange = datas[(int)ExportParameter.TARGET_RANGE].Value.ToString;
            this.CombineCount = datas[(int)ExportParameter.COMBINE_COUNT].Value.ToInt;
            CombinRanges = new List<(string SheetName, string Range, string Origin)>();

            //SourceSheetName 추가한 이유는 합칠때 원본이 되는 시트가 없으므로 시작 시트 지정용도로 사용하기 위함.
            // 추가되면서 추가 수정 필요한 부분 재확인 후 작업할것

            for (int i = (int)ExportParameter.COMBINE_RANGES; i < CombineCount + (int)ExportParameter.COMBINE_RANGES; ++i)
            {
                string[] val = datas[i].Value.ToString.Split('!');
                if (val.Length == 1)
                    CombinRanges.Add((null, val[0], datas[i].Value.ToString));
                else
                    CombinRanges.Add((val[0], val[1], datas[i].Value.ToString));
            }
        }

        /// <summary>
        /// 익스포트 세팅 시트에서 익스포트 세팅값들을 초기화 후 가져오기.
        /// </summary>
        /// <param name="_settingSheet"></param>
        /// <returns></returns>
        public static (List<ExportSetting> List, Dictionary<(int serverType, int LegionType, string SourceName), ExportSetting> Dic) GetExportSettings(Excel.Worksheet _settingSheet, string RangeValue)
        {
            List<ExportSetting> exportSettings = new List<ExportSetting>();
            Dictionary<(int serverType, int LegionType, string SourceName), ExportSetting> exportSettingDic = new Dictionary<(int, int, string), ExportSetting>();
            Range data = _settingSheet.Range[RangeValue];

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
                else if (cells[0].Type != null && cells[0].Type == typeof(string) && cells[0].Value.ToString == "Index")
                    continue;

                ExportSetting exportSetting = new ExportSetting(cells);
                exportSettings.Add(exportSetting);
                try
                {
                    exportSettingDic.Add(
                        (
                        serverType: exportSetting.ServerType,
                        LegionType: exportSetting.LegionType,
                        SourceName: exportSetting.SourceName
                        )
                        , exportSetting
                        );
                }
                catch (Exception e)
                {
                    Console.WriteLine($"ExportSetting 중복된 값이 있습니다. \n index ({row}) ServerType ({exportSetting.ServerType}) LegionType ({exportSetting.LegionType}) Name ({exportSetting.Name})");
                }

            }

            return (List: exportSettings, Dic: exportSettingDic);
        }
    }
}
