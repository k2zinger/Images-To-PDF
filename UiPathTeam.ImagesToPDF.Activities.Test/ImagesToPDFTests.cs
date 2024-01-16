using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Documents;

namespace UiPathTeam.ImagesToPDF.Activities.Test
{
    [TestClass]
    public class ImagesToPDFTests
    {
        [TestMethod]
        public void Test_ImageToPDF()
        {
            Assert.IsTrue(File.Exists(Execute_ImageToPDF()));
        }

        //[TestMethod]
        public void Test_ImageToPDF_in_loop()
        {
            List<string> images = new List<string>();
            for(int i = 0; i < 10; i++)
            {
                images.Add(Execute_ImageToPDF());
            }

            bool result = true;
            foreach(string image in images)
            {
                result = result && File.Exists(image);
            }

            Assert.IsTrue(result);
        }

        private string Execute_ImageToPDF()
        {
            var fileName = DateTimeOffset.Now.ToUnixTimeMilliseconds().ToString();
            var imagefile = CreateImage("iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAIAAACRXR/mAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAADHSURBVFhH7ZLRDYQgEAWty4Koh2psxmJwl+ydYC6nExP0482P8HaRyYapvBJpEaRFkBZBWgRpEaRFkBZBWgRpEaRFuKo1Tz+Y8xrlc5ZE2um02N8bHtCyMEhLRE1mac77tPeWP9zXWu3KiHxZb20yP+LZ4Gl9VdqNfw9jGaxlQU/YNHlNHtDq53IgGgZr9e9ozcmXS9qbrO6bU/2O+1qGXxlEsUk+UX1vxiU3qjUIaRGkRZAWQVoEaRGkRZAWQVoEaRFeqVXKBgryZeCzFyUAAAAAAElFTkSuQmCC", fileName);

            var imageToPDF = new ImageToPDF
            {
                ImageFile = new InArgument<String>((ctx) => imagefile),
                OutputFile = new InArgument<String>((ctx) => Path.Combine(Environment.CurrentDirectory, fileName + ".pdf")),
            };

            WorkflowInvoker.Invoke(imageToPDF);

            return Path.Combine(Environment.CurrentDirectory, fileName + ".pdf");
        }

        private string CreateImage(String Base64, String FileName)
        {
            byte[] bytes = Convert.FromBase64String(Base64);

            Image image;
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                image = Image.FromStream(ms);
            }

            string filePath = Path.Combine(Environment.CurrentDirectory, FileName + ".png");
            image.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);

            return filePath;
        }
    }
}
