using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using LLDOCX.CustomUserControl;
using LLDOCX.CustomEvents;
using LLDOCX.Helpers;

namespace LLDOCX
{
    public partial class ThisAddIn
    {
        private SettingUserControl _settingControl;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public MissingHeadingHandler MissingHeadingHandler { get; set; }
        public List<DocumentHeadingData> DocumentHeadingData { get; set; }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            MissingHeadingHandler = new MissingHeadingHandler();
            DocumentHeadingData = new List<DocumentHeadingData>();
        }
        public void ProcessSideBarPanel(bool isVisible)
        {
            if (_settingControl == null)
            {
                _settingControl = new SettingUserControl();
            }
            if (CustomTaskPanes.Count == 0)
            {
                myCustomTaskPane = this.CustomTaskPanes.Add(_settingControl, "LL-DOCX Setting", this.Application.ActiveWindow);
                myCustomTaskPane.Width = 450;
            }
            myCustomTaskPane.Visible = isVisible;
        }
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonExplorer();
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
