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
            string fileName = FilePath + "Add Image in Table.ppt";
            string img = FilePath + "image.jpg";
            
            Presentation MyPresentation = new Presentation(fileName);

            //Get First Slide
            ISlide sld = MyPresentation.Slides[0];

            //Creating a Bitmap Image object to hold the image file
            System.Drawing.Bitmap image = new Bitmap(img);

            //Create an IPPImage object using the bitmap object
            IPPImage imgx1 = MyPresentation.Images.AddImage(image);

            foreach (IShape shp in sld.Shapes)
                if (shp is ITable)
                {
                    ITable tbl = (ITable)shp;
                    
                    //Add image to first table cell
                    tbl[0, 0].FillFormat.FillType = FillType.Picture;
                    tbl[0, 0].FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
                    tbl[0, 0].FillFormat.PictureFillFormat.Picture.Image = imgx1;
                }
            //Save PPTX to Disk
            MyPresentation.Save(fileName, Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}
