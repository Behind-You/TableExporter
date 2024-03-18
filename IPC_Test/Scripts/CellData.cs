using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace IPC_Test
{
    public abstract class _ConvertableType
    {
        public Type Type;
        public object Value;

        public byte ToByte
        {
            get
            {
                if (Type == typeof(byte))
                {
                    return (byte)Value;
                }
                else if (Type == typeof(int))
                {
                    return (byte)Value;
                }
                else if (Type == typeof(double))
                {
                    return (byte)(double)Value;
                }
                else if (Type == typeof(string))
                {
                    return byte.Parse((string)Value);
                }
                else if (Type == typeof(float))
                {
                    return (byte)(float)Value;
                }
                else
                {
                    return 0;
                }
            }
        }

        public int ToInt
        {
            get
            {
                if (Type == typeof(byte))
                {
                    return (int)Value;
                }
                else if (Type == typeof(int))
                {
                    return (int)Value;
                }
                else if (Type == typeof(double))
                {
                    return (int)(double)Value;
                }
                else if (Type == typeof(string))
                {
                    return int.Parse(Value.ToString());
                }
                else if (Type == typeof(float))
                {
                    return (int)(float)Value;
                }
                else
                {
                    return 0;
                }
            }
        }

        //Value 값을 각 타입에 맞게 형변환 해서 전달하는 함수
        public long ToLong
        {
            get
            {
                if (Type == typeof(byte))
                {
                    return (long)Value;
                }
                else if (Type == typeof(int))
                {
                    return (long)Value;
                }
                else if (Type == typeof(double))
                {
                    return (long)Value;
                }
                else if (Type == typeof(string))
                {
                    return long.Parse(Value.ToString());
                }
                else if (Type == typeof(float))
                {
                    return (long)(float)Value;
                }
                else
                {
                    return 0;
                }
            }
        }

        public double ToDouble
        {
            get
            {
                if (Type == typeof(byte))
                {
                    return (double)Value;
                }
                else if (Type == typeof(int))
                {
                    return (double)Value;
                }
                else if (Type == typeof(double))
                {
                    return (double)Value;
                }
                else if (Type == typeof(string))
                {
                    return double.Parse(Value.ToString());
                }
                else if (Type == typeof(float))
                {
                    return (double)(float)Value;
                }
                else
                {
                    return 0;
                }
            }
        }

        public float ToFloat
        {
            get
            {
                if (Type == typeof(byte))
                {
                    return (float)Value;
                }
                else if (Type == typeof(int))
                {
                    return (float)Value;
                }
                else if (Type == typeof(double))
                {
                    return (float)Value;
                }
                else if (Type == typeof(string))
                {
                    return float.Parse(Value.ToString());
                }
                else if (Type == typeof(float))
                {
                    return (float)Value;
                }
                else
                {
                    return 0;
                }
            }
        }

        public new string ToString
        {
            get
            {
                if (typeof(string) == Type)
                {
                    return (string)Value;
                }
                else
                {
                    return Value.ToString();
                }
            }
        }

        public Type ToType
        {
            get
            {
                switch (ToString)
                {
                    case "System.Byte":
                    case "Byte":
                    case "byte":
                        return typeof(byte);
                    case "System.Int32":
                    case "Int32":
                    case "int32":
                    case "Int":
                    case "int":
                        return typeof(int);
                    case "System.Double":
                    case "Double":
                    case "double":
                        return typeof(double);
                    case "System.String":
                    case "String":
                    case "string":
                        return typeof(string);
                    case "System.Single":
                    case "Single":
                    case "single":
                    case "Float":
                    case "float":
                        return typeof(float);
                    default:
                        return null;
                }
            }
        }

    }

    public class CellValue : _ConvertableType
    {

    }


    /// <summary>
    /// 직접 제작한 Cell Data 클래스
    /// </summary>
    public class CellData
    {
        private readonly string _row;
        private readonly string _col;
        private readonly string _position;
        private readonly CellValue _cellValue;

        public string Row => _row;
        public string Col => _col;
        public string Position => _position;
        public CellValue Value => _cellValue;
        public Type Type => _cellValue.Type;


        public CellData(string position, object value)
        {
            var pos = position.Split('$');
            this._col = pos[1];
            this._row = pos[2];
            this._position = position;
            if (value == null)
            {
                this._cellValue = null;
                return;
            }
            this._cellValue = new CellValue();
            this._cellValue.Value = value;
            if (this._cellValue.Value == null)
                return;
            this._cellValue.Type = value.GetType();
        }





        /// <summary>
        /// Range를 CellData List로 변환
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static List<CellData> RangeToCellDataList(Range range, bool ISRemoveEmpty = false)
        {
            List<CellData> cells = new List<CellData>();
            foreach (Range cell in range.Cells)
            {
                CellData celldata = new CellData(cell.Address, cell.Value2);
                cells.Add(celldata);
                if (cell.Value2 == null)
                {
                    if (ISRemoveEmpty)
                    {
                        break;
                    }
                }
            }
            return cells;
        }

    }
}
