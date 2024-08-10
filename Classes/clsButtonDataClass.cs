using System.Windows.Media.Imaging;

namespace LifestyleDesign_r24.Classes
{
    internal class clsButtonDataClass
    {
        public PushButtonData Data { get; set; }
        public clsButtonDataClass(string name, string text, string className,
            Bitmap largeImage,
            Bitmap smallImage,
            string toolTip)
        {
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;

            Data.LargeImage = BitmapToImageSource(largeImage);
            Data.Image = BitmapToImageSource(smallImage);

            // set command availability
            Data.AvailabilityClassName = "LifestyleDesign_r24.Common.CommandAvailability";
        }
        public clsButtonDataClass(string name, string text, string className,
            Bitmap largeImage,
            Bitmap smallImage,
            Bitmap largeImageDark,
            Bitmap smallImageDark,
            string toolTip)
        {
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;

            Data.LargeImage = BitmapToImageSource(largeImage);
            Data.Image = BitmapToImageSource(smallImage);

            // set command availability
            Data.AvailabilityClassName = "LifestyleDesign_r24.Common.CommandAvailability";
        }
        public static Assembly GetAssembly()
        {
            return Assembly.GetExecutingAssembly();
        }
        public static string GetAssemblyName()
        {
            return GetAssembly().Location;
        }
        public static BitmapImage BitmapToImageSource(Bitmap bm)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }
    }
}
