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
            string fileName = FilePath + "AddPictureFrame.pptx";
            string imgFile = FilePath + "image.jpg";
            
            //Instantiate Prsentation class that represents the PPTX
            Presentation pres = new Presentation(fileName);

            //Get the first slide
            ISlide sld = pres.Slides[0];

            //Instantiate the ImageEx class
            Image img = (Image)new Bitmap(imgFile);
            IPPImage imgx = pres.Images.AddImage(img);

            //Add Picture Frame with height and width equivalent of Picture
            sld.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 150, imgx.Width, imgx.Height, imgx);
            pres.Save(fileName,Slides.Export.SaveFormat.Pptx);
        }
    }
}
