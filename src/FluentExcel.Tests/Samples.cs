using System;
using System.Collections.Generic;
using System.Text;
using Xunit;
using FluentExcel;
using System.IO;

namespace FluentExcel.Tests
{
    public class Samples
    {

        [Fact]
        public void ShouldGenerateSampleWorksheet()
        {
            List<SampleEntity> sampleData = new List<SampleEntity>() {
                new SampleEntity {Name = "User 1",BirthDate = new DateTime(1988,1,14),Debt = -5000 } ,
                new SampleEntity {Name = "User 2",BirthDate = new DateTime(1990,2,20),Debt = -150 },
                new SampleEntity {Name = "User 3",BirthDate = new DateTime(1995,9,23),Debt = 0}
            };

            sampleData.BuildExporter()
                                     .AddColumn(c => c.Name, "User Name")
                                     .AddColumn(c => c.BirthDate, "Birth Date", Format.Date)
                                     .AddColumn(c => c.Debt, "Total Debt", Format.Money)
                                     .WithWorksheetName("Planilha de Usuários")
                                     .WithHeaderStyle(new ColumnStyle()
                                     {
                                         BackgroundColor = System.Drawing.Color.Black,
                                         TextColor = System.Drawing.Color.White
                                     })
                                     .ExportToFile(new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "Sample-worksheet.xlsx")));
        }

    }
}
