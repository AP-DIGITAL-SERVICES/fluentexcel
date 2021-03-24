using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;

namespace FluentExcel
{
    internal sealed class ColumnSettings<T> where T : class
    {
        public int Order { get; set; }
        public Expression<Func<T, object>> Property { get; set; }
        public string DisplayName { get; set; }

        public Format Format { get; set; }

        public ColumnStyle Style { get; set; }

        public string TextFormat { get; set; }

    }
}
