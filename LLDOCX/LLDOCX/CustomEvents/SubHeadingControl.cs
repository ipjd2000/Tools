using LLDOCX.Helpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace LLDOCX.CustomEvents
{
    public partial class SubHeadingControl : UserControl
    {
        private readonly DocumentHeadingData _documentHeadingData;
        public SubHeadingControl(DocumentHeadingData documentHeadingData)
        {
            InitializeComponent();
            _documentHeadingData = documentHeadingData;
            this.Load += OnSubHeadingControlLoad;
        }
        private void OnSubHeadingControlLoad(object sender, EventArgs e)
        {
            label1.Text = _documentHeadingData.Heading;
            comboxSubHeading.DataSource = _documentHeadingData.SubHeadings;
        }

        private void OncomboxSubHeadingSelectionChangeCommitted(object sender, EventArgs e)
        {
            string selectedHeading = (string)comboxSubHeading.SelectedItem;
            var app = Globals.ThisAddIn.Application;
            Word.Range selectedRang = app.Selection.Range;
            if (!string.IsNullOrEmpty(selectedRang.Text))
            {
                selectedRang.Text = selectedHeading + Environment.NewLine;
                selectedRang.Underline = Word.WdUnderline.wdUnderlineNone;
            }
        }
    }
}
