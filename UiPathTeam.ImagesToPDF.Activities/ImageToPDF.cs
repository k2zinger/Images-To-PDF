using System;
using System.Activities;
using System.ComponentModel;

namespace UiPathTeam.ImagesToPDF.Activities
{
    [DisplayName("Image To PDF"), Description("Convert an image to a PDF document")]
    public class ImageToPDF : ImagesToPDF
    {
        #region Properties

        [Browsable(false)]
        public override InArgument<String> ImagesFolder { get; set; }

        #endregion

    }
}