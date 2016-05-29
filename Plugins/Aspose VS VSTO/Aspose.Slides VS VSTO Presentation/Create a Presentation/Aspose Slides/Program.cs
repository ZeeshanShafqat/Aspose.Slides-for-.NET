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
            string fileName = FilePath + "CreatePresentation.pptx";
            //Create a presentation
            Presentation pres = new Presentation(fileName);

            //Add the title slide
            ISlide slide = pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);

            //Set the title text
            ((IAutoShape)slide.Shapes[0]).TextFrame.Text = "Slide Title Heading";

            //Set the sub title text
            ((IAutoShape)slide.Shapes[1]).TextFrame.Text = "Slide Title Sub-Heading";

            //Write output to disk
            pres.Save(fileName,Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}
