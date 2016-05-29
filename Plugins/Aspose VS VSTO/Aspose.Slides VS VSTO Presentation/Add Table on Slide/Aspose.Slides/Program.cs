using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Drawing;
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
            string fileName = FilePath + "AddTableonSlide.pptx";
            //Instantiate Prsentation class that represents the PPTX
            Presentation pres = new Presentation(fileName);

            //Get the first slide
            ISlide sld = pres.Slides[0];

            //Define columns with widths and rows with heights
            double[] dblCols = { 50, 50, 50 };
            double[] dblRows = { 50, 30, 30, 30, 30 };

            //Add table shape to slide
            ITable tbl = sld.Shapes.AddTable(100, 50, dblCols, dblRows);

            //Set border format for each cell
            foreach (IRow row in tbl.Rows)
                foreach (ICell cell in row)
                {
                    cell.BorderTop.FillFormat.FillType = FillType.Solid;
                    cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderTop.Width = 5;

                    cell.BorderBottom.FillFormat.FillType = FillType.Solid;
                    cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderBottom.Width = 5;

                    cell.BorderLeft.FillFormat.FillType = FillType.Solid;
                    cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderLeft.Width = 5;

                    cell.BorderRight.FillFormat.FillType = FillType.Solid;
                    cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderRight.Width = 5;
                }

            //Merge cells 1 & 2 of row 1
            tbl.MergeCells(tbl[0, 0], tbl[1, 0], false);

            //Add text to the merged cell
            tbl[0, 0].TextFrame.Text = "Merged Cells";
            pres.Save(fileName,Slides.Export.SaveFormat.Pptx);
        }
    }
}
