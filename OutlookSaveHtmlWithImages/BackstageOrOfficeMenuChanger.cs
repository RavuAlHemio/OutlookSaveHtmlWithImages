using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OutlookSaveHtmlWithImages.Properties;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.


namespace OutlookSaveHtmlWithImages
{
    [ComVisible(true)]
    public class BackstageOrOfficeMenuChanger : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private bool? _hasBackstage;

        public BackstageOrOfficeMenuChanger()
        {
            _hasBackstage = null;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            if (!_hasBackstage.HasValue)
            {
                int officeMajorVersion = int.Parse(Globals.ThisAddIn.Application.Version.Split('.')[0]);
                _hasBackstage = (officeMajorVersion > 12);
            }
            return _hasBackstage.Value
                ? GetResourceText("OutlookSaveHtmlWithImages.BackstageChanger.xml")
                : GetResourceText("OutlookSaveHtmlWithImages.OfficeMenuChanger.xml");
        }

        #endregion

        #region Ribbon Callbacks
        public void SaveHtmlWithImages(Office.IRibbonControl control)
        {
            var inspector = control.Context as Outlook.Inspector;
            if (inspector == null)
            {
                return;
            }

            var mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem == null)
            {
                return;
            }

            if (mailItem.BodyFormat != Outlook.OlBodyFormat.olFormatHTML)
            {
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                AddExtension = true,
                AutoUpgradeEnabled = true,
                OverwritePrompt = true,
                ValidateNames = true,
                Filter = Resources.HtmlFileSaveFileDialogFilter
            };
            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            mailItem.SaveAs(saveFileDialog.FileName, Outlook.OlSaveAsType.olHTML);
            HtmlSaver.PostprocessHtml(saveFileDialog.FileName);
        }

        public bool IsVisible(Office.IRibbonControl control)
        {
            var inspector = control.Context as Outlook.Inspector;
            if (inspector == null)
            {
                return false;
            }

            var mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem == null)
            {
                return false;
            }

            return (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML);
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "saveHtmlWithImagesBtn":
                    return Properties.Resources.SaveHtmlWithImagesBtnLabel;
                default:
                    return null;
            }
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this._ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
