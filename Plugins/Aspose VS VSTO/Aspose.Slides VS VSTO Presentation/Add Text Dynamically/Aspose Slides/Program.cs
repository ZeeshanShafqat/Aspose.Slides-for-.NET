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
            string fileName = FilePath + "AddTextDynamically.pptx";
            //Create a presentation
            Presentation pres = new Presentation(fileName);

            //Blank slide is added by default, when you create
            //presentation from default constructor
            //So, we don't need to add any blank slide
            ISlide sld = pres.Slides[1];

            //Add a textbox
            //To add it, we will first add a rectangle
            IShape shp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);

            //Hide its line
            shp.LineFormat.Style = LineStyle.NotDefined;

            //Then add a textframe inside it
            ITextFrame tf = ((IAutoShape)shp).TextFrame;

            //Set a text
            tf.Text = "Text added dynamically";
            IPortion port = tf.Paragraphs[0].Portions[0];

            port.PortionFormat.FontBold = NullableBool.True;
            port.PortionFormat.FontHeight = 32;


            //Write the output to disk
            pres.Save(fileName, Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}
