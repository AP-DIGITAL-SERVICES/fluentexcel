using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Text;

namespace FluentExcel
{
    public static class ExcelExporter
    {
        public static IExportedFileBuilder<T> BuildExporter<T>(this IEnumerable<T> content) where T : class
        {
            return new ExportedFileBuilder<T>(content);
        }
    }
}
