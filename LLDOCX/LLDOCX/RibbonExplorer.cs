using LLDOCX.Options;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;


namespace LLDOCX
{
    [ComVisible(true)]
    public class RibbonExplorer : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private bool isPressed;
        public RibbonExplorer()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("LLDOCX.RibbonExplorer.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public Bitmap GetImages(Office.IRibbonControl control)
        {
            if (control.Id == "btnScan")
                return Properties.Resources.LL_logo;
            if (isPressed)
                return Properties.Resources.toggle_on;
            return Properties.Resources.toggle_off;
        }
        public void OnScanning(Office.IRibbonControl control)
        {
            var option = ProcessData.GetData();
            var missingheadings = Helpers.WordDocumentHelper.ValidateDocument(option);
            if (missingheadings.Count > 0 || Globals.ThisAddIn.DocumentHeadingData.Count > 0)
            {
                if (!this.isPressed)
                {
                    this.isPressed = true;
                    InvalidateControl(control.Id);
                }
                Globals.ThisAddIn.ProcessSideBarPanel(this.isPressed);
                Globals.ThisAddIn.MissingHeadingHandler.OnEmailItemChanged(missingheadings);
            }
        }
        public bool GetPressed(Office.IRibbonControl control)
        {
            return isPressed;
        }
        public void InvalidateControl(string controlId)
        {
            ribbon.InvalidateControl(controlId);
        }
        public void ShowSetting(Office.IRibbonControl control, bool isPressed)
        {
            this.isPressed = isPressed;
            InvalidateControl(control.Id);
            Globals.ThisAddIn.ProcessSideBarPanel(this.isPressed);
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
