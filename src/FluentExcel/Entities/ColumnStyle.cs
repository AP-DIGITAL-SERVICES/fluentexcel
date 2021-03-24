using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace FluentExcel
{
    public class ColumnStyle
    {
        public Color TextColor { get; set; }
        public Color BackgroundColor { get; set; }
        public bool IsBold { get; set; }
        public Color BorderColor { get; set; }
        public bool SetBorder { get; set; }
        public float FontSize { get; set; } = 10;
        public OfficeOpenXml.Style.ExcelHorizontalAlignment HorizontalAligment { get; set; } = OfficeOpenXml.Style.ExcelHorizontalAlignment.General;
        public OfficeOpenXml.Style.ExcelVerticalAlignment VerticalAligment { get; set; } = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
    }
}
