using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Slides;
using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "CreateandEmbedExcelChart.pptx";
            //Create a workbook
            Workbook wb = new Workbook();

            //Add an excel chart
            int chartSheetIndex = AddExcelChartInWorkbook(wb);

            wb.Worksheets.SetOleSize(0, 5, 0, 5);

            Bitmap imgChart = wb.Worksheets[chartSheetIndex].Charts[0].ToImage();

            //Save the workbook to stream
            MemoryStream wbStream = wb.SaveToStream();

            //Create a presentation            
            Presentation pres = new Presentation();
            ISlide sld = pres.Slides[0];

            //Add the workbook on slide
            AddExcelChartInPresentation(pres, sld, wbStream, imgChart);

            //Write the output presentation on disk
            pres.Save(fileName, Aspose.Slides.Export.SaveFormat.Pptx);
        }

        static int AddExcelChartInWorkbook(Workbook wb)
        {
            //Add a new worksheet to populate cells with data
            int dataSheetIdx = wb.Worksheets.Add();

            Worksheet dataSheet = wb.Worksheets[dataSheetIdx];

            string sheetName = "DataSheet";

            dataSheet.Name = sheetName;

            //Populate DataSheet with data
            dataSheet.Cells["A2"].PutValue("N. America");
            dataSheet.Cells["A3"].PutValue("S. America");
            dataSheet.Cells["A4"].PutValue("Europe");
            dataSheet.Cells["A5"].PutValue("Asia");

            dataSheet.Cells["B1"].PutValue("Q1");
            dataSheet.Cells["B2"].PutValue(1.5);
            dataSheet.Cells["B3"].PutValue(2);
            dataSheet.Cells["B4"].PutValue(2.25);
            dataSheet.Cells["B5"].PutValue(2.5);

            dataSheet.Cells["C1"].PutValue("Q2");
            dataSheet.Cells["C2"].PutValue(2);
            dataSheet.Cells["C3"].PutValue(1.75);
            dataSheet.Cells["C4"].PutValue(2);
            dataSheet.Cells["C5"].PutValue(2.5);

            dataSheet.Cells["D1"].PutValue("Q3");
            dataSheet.Cells["D2"].PutValue(1.5);
            dataSheet.Cells["D3"].PutValue(2);
            dataSheet.Cells["D4"].PutValue(2.5);
            dataSheet.Cells["D5"].PutValue(2);

            dataSheet.Cells["E1"].PutValue("Q4");
            dataSheet.Cells["E2"].PutValue(2.5);
            dataSheet.Cells["E3"].PutValue(2);
            dataSheet.Cells["E4"].PutValue(2);
            dataSheet.Cells["E5"].PutValue(2.75);

            //Add a chart sheet
            int chartSheetIdx = wb.Worksheets.Add(SheetType.Chart);

            Worksheet chartSheet = wb.Worksheets[chartSheetIdx];

            chartSheet.Name = "ChartSheet";

            //Add a chart in ChartSheet with data series from DataSheet

            int chartIdx = chartSheet.Charts.Add(ChartType.Column3DClustered, 0, 5, 0, 5);

            Aspose.Cells.Charts.Chart chart = chartSheet.Charts[chartIdx];

            chart.NSeries.Add(sheetName + "!A1:E5", false);

            //Setting Chart's Title
            chart.Title.Text = "Sales by Quarter";

            //Setting the foreground color of the plot area
            chart.PlotArea.Area.ForegroundColor = Color.White;

            //Setting the background color of the plot area
            chart.PlotArea.Area.BackgroundColor = Color.White;

            //Setting the foreground color of the chart area
            chart.ChartArea.Area.BackgroundColor = Color.White;

            chart.Title.TextFont.Size = 16;

            //Setting the title of category axis of the chart
            chart.CategoryAxis.Title.Text = "Fiscal Quarter";

            //Setting the title of value axis of the chart
            chart.ValueAxis.Title.Text = "Billions";

            //Set ChartSheet an active sheet
            wb.Worksheets.ActiveSheetIndex = chartSheetIdx;

            return chartSheetIdx;
        }

        private static void AddExcelChartInPresentation(Presentation pres, ISlide sld, Stream wbStream, Bitmap imgChart)
        {
            float oleWidth = pres.SlideSize.Size.Width;
            float oleHeight = pres.SlideSize.Size.Height;
            int x = 0;
            byte[] chartOleData = new byte[wbStream.Length];
            wbStream.Position = 0;
            wbStream.Read(chartOleData, 0, chartOleData.Length);
            IOleObjectFrame oof = null;
            oof = sld.Shapes.AddOleObjectFrame(x, 0, oleWidth, oleHeight, "Excel.Sheet.8", chartOleData);
            oof.SubstitutePictureFormat.Picture.Image = pres.Images.AddImage((System.Drawing.Image)imgChart);
        }
    }
}
