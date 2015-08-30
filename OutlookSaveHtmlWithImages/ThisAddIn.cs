using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookSaveHtmlWithImages
{
    public partial class ThisAddIn
    {
        /// <remarks>Cached to make sure our event handlers are not garbage-collected.</remarks>
        private Outlook.Inspectors _inspectors;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += InspectorsOnNewInspector;
        }

        private void InspectorsOnNewInspector(Outlook.Inspector inspector)
        {
            //throw new NotImplementedException();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new BackstageChanger();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
        }
        
        #endregion
    }
}
