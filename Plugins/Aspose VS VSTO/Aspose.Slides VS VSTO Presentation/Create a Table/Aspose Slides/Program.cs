using Aspose.Slides;
using System;
using System.Collections.Generic;
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
            string fileName = FilePath + "CreateTable.pptx";
            Presentation pres = new Presentation(fileName);

            //Access first slide
            ISlide sld = pres.Slides[0];

            //Define columns with widths and rows with heights
            double[] dblCols = { 50, 50, 50 };
            double[] dblRows = { 50, 30, 30, 30, 30 };

            //Add a table
            Aspose.Slides.ITable tbl = sld.Shapes.AddTable(50, 50, dblCols, dblRows);

            //Set border format for each cell
            foreach (IRow row in tbl.Rows)
            {
                foreach (ICell cell in row)
                {

                    //Get text frame of each cell
                    ITextFrame tf = cell.TextFrame;
                    //Add some text
                    tf.Text = "T" + cell.FirstRowIndex.ToString() + cell.FirstColumnIndex.ToString();
                    //Set font size of 10
                    tf.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;
                    tf.Paragraphs[0].ParagraphFormat.Bullet.Type = BulletType.None;
                }
            }

            //Write the presentation to the disk
            pres.Save(fileName,Slides.Export.SaveFormat.Pptx);
        }
    }
}
