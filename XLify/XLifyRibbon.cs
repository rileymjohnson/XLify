using System;
using Office = Microsoft.Office.Core;
using System.Drawing;
using XLify.Properties;
using System.Runtime.InteropServices;

namespace XLify
{
    [ComVisible(true)]
    public class XLifyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                // Ribbon XML: Add a large toggle button on the built-in Add-Ins tab, using custom image resources
                return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' loadImage='LoadImage'>
  <ribbon>
    <tabs>
      <tab idMso='TabAddIns'>
        <group id='grpXLifyMain' label='XLify'>
          <toggleButton id='btnOpenTaskPane' label='Open' size='large' onAction='OnTogglePane' getPressed='GetPanePressed' image='XLify_logo_32' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
            }
            catch
            {
                // If anything goes wrong, return no UI to avoid Excel disabling the add-in
                return null;
            }
        }
        public void OnTogglePane(Office.IRibbonControl control, bool pressed)
        {
            // Ignore 'pressed' and toggle based on actual pane state to avoid UI desync
            try { Globals.ThisAddIn?.ToggleTaskPane(); } catch { }
        }

        public bool GetPanePressed(Office.IRibbonControl control)
        {
            try { return Globals.ThisAddIn?.IsTaskPaneVisible() == true; } catch { return false; }
        }

        // Callback for ribbon image loading (loadImage="LoadImage")
        public stdole.IPictureDisp LoadImage(string imageId)
        {
            try
            {
                switch (imageId)
                {
                    case "XLify_logo_32":
                        return ImageConverter.ToIPictureDisp(Resources.XLify_logo_32);
                    case "XLify_logo_16":
                        return ImageConverter.ToIPictureDisp(Resources.XLify_logo_16);
                    default:
                        return null;
                }
            }
            catch { return null; }
        }

        internal static Office.IRibbonUI RibbonUI { get; private set; }

        public void OnLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            RibbonUI = ribbonUI;
        }

        internal static void Invalidate(string controlId)
        {
            try { RibbonUI?.InvalidateControl(controlId); } catch { }
        }

        // Helper to convert Bitmap -> IPictureDisp for Ribbon images
        private static class ImageConverter
        {
            private class AxHostConverter : System.Windows.Forms.AxHost
            {
                private AxHostConverter() : base(null) { }
                public static stdole.IPictureDisp ToIPictureDisp(Image image) => (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }

            public static stdole.IPictureDisp ToIPictureDisp(Bitmap bmp)
            {
                return AxHostConverter.ToIPictureDisp(bmp);
            }
        }
    }
}
