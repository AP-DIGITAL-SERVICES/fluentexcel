using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Text;

namespace FluentExcel
{
    public interface IExportedFileBuilder<T> where T : class
    {
        IExportedFileBuilder<T> AddColumn(Expression<Func<T, object>> column, string displayName, Format format = Format.Raw);
        IExportedFileBuilder<T> WithWorksheetName(string workSheetName);
        IExportedFileBuilder<T> WithDefaultColumnStyle(ColumnStyle style);
        IExportedFileBuilder<T> WithHeaderStyle(ColumnStyle style);
        Stream ExportToStream();
        void ExportToFile(FileInfo file);
    }
}
