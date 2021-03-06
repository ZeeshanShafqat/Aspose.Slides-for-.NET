﻿using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSVSTO
{
    class Program
    {
        static void Main(string[] args)
        {
            //Instantiate the Presentation class that represents the presentation file
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "SetBackgroungColour.pptx";
            using (Presentation pres = new Presentation(fileName))
            {

                //Set the background color of the Master ISlide to Forest Green
                
                pres.Masters[0].Background.Type = BackgroundType.OwnBackground;
                pres.Masters[0].Background.FillFormat.FillType = FillType.Solid;
                pres.Masters[0].Background.FillFormat.SolidFillColor.Color = Color.ForestGreen;

                //Write the presentation to disk
                pres.Save(fileName, SaveFormat.Pptx);

            }
 
        }
    }
}
