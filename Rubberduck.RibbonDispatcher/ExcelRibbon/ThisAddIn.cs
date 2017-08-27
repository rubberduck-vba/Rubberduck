using System;

using Microsoft.Office.Tools.Excel;

using System.Windows.Forms;

namespace RubberDuck.RibbonDispatcher {
    using Office = Microsoft.Office.Core;

    public partial class ThisAddIn {
        private void ThisAddIn_Startup(object sender, EventArgs e) {
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) {
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() => new CustomAppRibbon();

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
