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
            string fileName = FilePath + "FormatText.pptx";
            string destfileName = FilePath + "FormatText-Output.pptx";
            //Open the presentation
            Presentation pres = new Presentation(fileName);

            //Access the first slide
            ISlide slide = pres.Slides[0];

            //Access the third shape
            IShape shp = slide.Shapes[1];

            //Change its text's font to Verdana and height to 32
            ITextFrame tf = ((IAutoShape)shp).TextFrame;
            IParagraph para = tf.Paragraphs[0];
            IPortion port = para.Portions[0];
            port.PortionFormat.LatinFont = new FontData("Verdana");

            port.PortionFormat.FontHeight = 32;

            //Bolden it
            port.PortionFormat.FontBold = NullableBool.True;

            //Italicize it
            port.PortionFormat.FontItalic = NullableBool.True;

            //Change text color
            //Set font color
            port.PortionFormat.FillFormat.FillType = FillType.Solid;
            port.PortionFormat.FillFormat.SolidFillColor.Color = Color.FromArgb(0x33, 0x33, 0xCC);

            //Change shape background color
            shp.FillFormat.FillType = FillType.Solid;
            shp.FillFormat.SolidFillColor.Color = Color.FromArgb(0xCC, 0xCC, 0xFF);

            //Write the output to disk
            pres.Save(destfileName,Slides.Export.SaveFormat.Pptx);
        }
    }
}
