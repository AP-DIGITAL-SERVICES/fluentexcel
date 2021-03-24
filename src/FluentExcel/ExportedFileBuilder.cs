using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Text;
using System.Linq;
using System.Drawing;
using System.Reflection;

namespace FluentExcel
{
    public class ExportedFileBuilder<T> : IExportedFileBuilder<T> where T : class
    {
        #region .: Variables :.
        private IEnumerable<T> _content;
        private ICollection<ColumnSettings<T>> _propertyMapping = new HashSet<ColumnSettings<T>>();
        private string _workSheetName = string.Empty;
        private bool _useDefaultColumnStyle = true;

        private ColumnStyle _headerStyle = new ColumnStyle
        {
            BackgroundColor = Color.Transparent,
            TextColor = Color.Black,
            BorderColor = Color.Transparent,
            IsBold = true,
            SetBorder = false
        };

        private ColumnStyle _defaultColumnStyle = new ColumnStyle
        {
            BackgroundColor = Color.Transparent,
            TextColor = Color.Black,
            BorderColor = Color.Transparent,
            IsBold = false,
            SetBorder = false
        };
        #endregion

        #region .: Constructor :.
        public ExportedFileBuilder(IEnumerable<T> content)
        {
            _content = content;

        }
        #endregion

        #region .: Public API :.
        public IExportedFileBuilder<T> AddColumn(Expression<Func<T, object>> column, string displayName, Format format = Format.Raw)
        {
            _propertyMapping.Add(new ColumnSettings<T>
            {
                DisplayName = displayName,
                Property = column,
                Format = format
            });

            return this;
        }
        private ExcelPackage BuildExcelPackage(Func<ExcelPackage> action)
        {
            return action();
        }

        public Stream ExportToStream()
        {
            Stream fileStream = new MemoryStream();

            Func<ExcelPackage> loadExcelPackage = () => new ExcelPackage();

            BuildWorkSheet(loadExcelPackage, c =>
            {
                c.SaveAs(fileStream);
            });

            return fileStream;
        }
        public void ExportToFile(FileInfo file)
        {

            Func<ExcelPackage> loadExcelPackage = () => new ExcelPackage();

            BuildWorkSheet(loadExcelPackage, c =>
            {
                c.SaveAs(file);
            });
        }

        public IExportedFileBuilder<T> WithDefaultColumnStyle(ColumnStyle style)
        {
            _defaultColumnStyle = style;
            _useDefaultColumnStyle = true;

            return this;
        }

        public IExportedFileBuilder<T> WithHeaderStyle(ColumnStyle settings)
        {
            _headerStyle = settings;

            return this;
        }

        public IExportedFileBuilder<T> WithWorksheetName(string workSheetName)
        {
            _workSheetName = workSheetName;

            return this;
        }
        #endregion

        #region .: Privated Methods :.
        private void BuildWorkSheet(Func<ExcelPackage> loadExcelPackage, Action<ExcelPackage> saveAction)
        {
            using (ExcelPackage package = BuildExcelPackage(loadExcelPackage))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(_workSheetName);

                int columnIndex = 1;

                foreach (ColumnSettings<T> propertySettings in _propertyMapping.OrderBy(c => c.Order))
                {
                    worksheet.Cells[1, columnIndex].Value = propertySettings.DisplayName;

                    worksheet.Cells[1, columnIndex].Style.Font.Size = _headerStyle.FontSize;
                    worksheet.Cells[1, columnIndex].Style.HorizontalAlignment = _headerStyle.HorizontalAligment;
                    worksheet.Cells[1, columnIndex].Style.VerticalAlignment = _headerStyle.VerticalAligment;
                    worksheet.Cells[1, columnIndex].Style.Font.Bold = _headerStyle.IsBold;
                    worksheet.Cells[1, columnIndex].Style.Font.Color.SetColor(_headerStyle.TextColor);
                    worksheet.Cells[1, columnIndex].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[1, columnIndex].Style.Fill.BackgroundColor.SetColor(_headerStyle.BackgroundColor);
                    worksheet.Cells[1, columnIndex].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, _headerStyle.BorderColor);

                    columnIndex++;

                }

                Dictionary<ColumnSettings<T>, PropertyInfo> propertyMappingCache = new Dictionary<ColumnSettings<T>, PropertyInfo>();

                foreach (ColumnSettings<T> column in _propertyMapping.OrderBy(c => c.Order))
                {
                    propertyMappingCache[column] = GetPropertyFromExpression(column.Property);
                }

                int rowIndex = 2;
                columnIndex = 1;

                foreach (T row in _content)
                {
                    foreach (KeyValuePair<ColumnSettings<T>, PropertyInfo> column in propertyMappingCache)
                    {
                        PropertyInfo property = column.Value;

                        worksheet.Cells[rowIndex, columnIndex].Value = property.GetValue(row);

                        ColumnStyle style = column.Key.Style != null && !_useDefaultColumnStyle ? column.Key.Style : _defaultColumnStyle;


                        worksheet.Cells[rowIndex, columnIndex].Style.Font.Size = style.FontSize;
                        worksheet.Cells[rowIndex, columnIndex].Style.HorizontalAlignment = style.HorizontalAligment;
                        worksheet.Cells[rowIndex, columnIndex].Style.VerticalAlignment = style.VerticalAligment;
                        worksheet.Cells[rowIndex, columnIndex].Style.Font.Bold = style.IsBold;
                        worksheet.Cells[rowIndex, columnIndex].Style.Font.Color.SetColor(style.TextColor);
                        worksheet.Cells[rowIndex, columnIndex].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[rowIndex, columnIndex].Style.Fill.BackgroundColor.SetColor(style.BackgroundColor);
                        worksheet.Cells[rowIndex, columnIndex].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, style.BorderColor);
                        worksheet.Cells[rowIndex, columnIndex].Style.Numberformat.Format = column.Key.TextFormat;

                        columnIndex++;

                    }

                    columnIndex = 1;
                    rowIndex++;
                }

                worksheet.Cells.AutoFitColumns(0);

                package.Save();

                saveAction(package);

            }
        }

        private PropertyInfo GetPropertyFromExpression(Expression<Func<T, object>> GetPropertyLambda)
        {
            MemberExpression Exp = null;

            if (GetPropertyLambda.Body is UnaryExpression)
            {
                var UnExp = (UnaryExpression)GetPropertyLambda.Body;
                if (UnExp.Operand is MemberExpression)
                {
                    Exp = (MemberExpression)UnExp.Operand;
                }
                else
                    throw new ArgumentException();
            }
            else if (GetPropertyLambda.Body is MemberExpression)
            {
                Exp = (MemberExpression)GetPropertyLambda.Body;
            }
            else
            {
                throw new ArgumentException();
            }

            return (PropertyInfo)Exp.Member;
        }
        #endregion
    }
}
