namespace FuniglooExportData
{
    public static class SettingTypes
    {
        public enum ExportParameter
        {
            INDEX = 1,
            SETTING_NAME,
            NAME,
            PATH,
            SOURCE_RANGE,
            TARGET_RANGE,
            COMBINE_COUNT,
            COMBINE_RANGES,
        }

        public static int ToInt(this ExportParameter type)
        {
            switch (type)
            {
                case ExportParameter.INDEX:
                    return 1;
                case ExportParameter.SETTING_NAME:
                    return 2;
                case ExportParameter.NAME:
                    return 3;
                case ExportParameter.PATH:
                    return 4;
                case ExportParameter.SOURCE_RANGE:
                    return 5;
                case ExportParameter.TARGET_RANGE:
                    return 6;
                case ExportParameter.COMBINE_COUNT:
                    return 7;
                case ExportParameter.COMBINE_RANGES:
                    return 8;
                default:
                    return -1;
            }
        }
        public static string ToToken(this ExportParameter type)
        {
            switch (type)
            {
                case ExportParameter.INDEX:
                    return "A";
                case ExportParameter.SETTING_NAME:
                    return "B";
                case ExportParameter.NAME:
                    return "C";
                case ExportParameter.PATH:
                    return "D";
                case ExportParameter.SOURCE_RANGE:
                    return "E";
                case ExportParameter.TARGET_RANGE:
                    return "F";
                case ExportParameter.COMBINE_COUNT:
                    return "G";
                case ExportParameter.COMBINE_RANGES:
                    return "H";
                default:
                    return "";
            }
        }
        public static string ToText(this ExportParameter type)
        {
            switch (type)
            {
                case ExportParameter.INDEX:
                    return "Index";
                case ExportParameter.SETTING_NAME:
                    return "SettingName";
                case ExportParameter.NAME:
                    return "Name";
                case ExportParameter.PATH:
                    return "Path";
                case ExportParameter.SOURCE_RANGE:
                    return "SourceRange";
                case ExportParameter.TARGET_RANGE:
                    return "TargetRange";
                case ExportParameter.COMBINE_COUNT:
                    return "CombineCount";
                case ExportParameter.COMBINE_RANGES:
                    return "CombineRange{0}";
                default:
                    return "";
            }
        }
    }

}
