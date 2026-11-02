using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

using LLDOCX.Options;
using LLDOCX.Helpers;
using LLDOCX.CustomEvents;

namespace LLDOCX.CustomUserControl
{
    public partial class SettingUserControl : UserControl
    {
        private SettingOption _option;
        public SettingUserControl()
        {
            InitializeComponent();
            Globals.ThisAddIn.MissingHeadingHandler.OnUpdateMissingHeadingData += OnUpdateMissingHeadingData;
            this.Load += OnSettingUserControlLoaded;
        }
        private void OnUpdateMissingHeadingData(object sender, List<DocumentHeadingData> missingHeadings)
        {
            Globals.ThisAddIn.DocumentHeadingData = missingHeadings.OrderBy(x => x.Index).ToList();
            LoadMissingHeading(missingHeadings.Count > 0);
        }
        private void LoadMissingHeading(bool isvisible)
        {
            panel1.Controls.Clear();
            panel1.Visible = isvisible;
            var groupbox = new GroupBox();
            groupbox.Dock = DockStyle.Fill;
            groupbox.Text = "Missing Headings";
            panel1.Controls.Add(groupbox);
            foreach (var heading in Globals.ThisAddIn.DocumentHeadingData)
            {
                var headingControl = new SubHeadingControl(heading);
                headingControl.Dock = DockStyle.Top;
                groupbox.Controls.Add(headingControl);
            }
        }

        private void OnSettingUserControlLoaded(object sender, EventArgs e)
        {
            _option = ProcessData.GetData();
            LoadSaveData();
            LoadMissingHeading(false);
        }
        private void LoadSaveData()
        {
            ckbPageSize.Checked = _option.CheckPageSize;
            ckbPageOrientation.Checked = _option.CheckPageOrientation;
            ckbMarkup.Checked = _option.CheckMarkup;
            ckbKeepTracking.Checked = _option.KeepTrackingOn;
            ckbFont.Checked = _option.CheckFont;
            ckbhyperlink.Checked = _option.CheckHyperlink;
            ckbComment.Checked = _option.CheckComment;
            ckbBookmarks.Checked = _option.CheckBookmarker;
            ckbDocProperty.Checked = _option.RemoveDocumentProperty;
            cbkHeadingValidation.Checked = _option.HeadingValidation;
            ckbCheckAll.Checked = _option.CheckedAll;
        }

        private void OnButnResetClicked(object sender, EventArgs e)
        {
            _option = new SettingOption();
            ProcessData.SaveData(_option);
            LoadSaveData();
            MessageBox.Show("Reset Setting successful", LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OnBtnSaveChangeClicked(object sender, EventArgs e)
        {
            SaveAllSettng();
            MessageBox.Show("Saved Setting successful", LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ckbCheckAll_CheckedChanged(object sender, EventArgs e)
        {
            ckbPageSize.Checked = ckbCheckAll.Checked;
            ckbPageOrientation.Checked = ckbCheckAll.Checked;
            ckbMarkup.Checked = ckbCheckAll.Checked;
            ckbKeepTracking.Checked = ckbCheckAll.Checked;
            ckbFont.Checked = ckbCheckAll.Checked;
            ckbhyperlink.Checked = ckbCheckAll.Checked;
            ckbComment.Checked = ckbCheckAll.Checked;
            ckbBookmarks.Checked = ckbCheckAll.Checked;
            ckbDocProperty.Checked = ckbCheckAll.Checked;
            cbkHeadingValidation.Checked = ckbCheckAll.Checked;
            SaveAllSettng();
        }
        private void SaveAllSettng()
        {
            _option.CheckPageSize = ckbPageSize.Checked;
            _option.CheckPageOrientation = ckbPageOrientation.Checked;
            _option.CheckMarkup = ckbMarkup.Checked;
            _option.KeepTrackingOn = ckbKeepTracking.Checked;
            _option.CheckFont = ckbFont.Checked;
            _option.CheckHyperlink = ckbhyperlink.Checked;
            _option.CheckComment = ckbComment.Checked;
            _option.CheckBookmarker = ckbBookmarks.Checked;
            _option.RemoveDocumentProperty = ckbDocProperty.Checked;
            _option.HeadingValidation = cbkHeadingValidation.Checked;
            _option.CheckedAll = ckbCheckAll.Checked;
            ProcessData.SaveData(_option);
        }

        private void ckbFont_CheckedChanged(object sender, EventArgs e)
        {
            btnFontName.Enabled = ckbFont.Checked;
        }

        private void btnFontName_Click(object sender, EventArgs e)
        {
            using (var customFont = new CustomUserControl.CustomFontDialog(_option))
            {
                if (customFont.ShowDialog() == DialogResult.OK)
                {
                    _option.DefaultFontName = customFont.SelectedFont;
                    SaveAllSettng();
                }
            }
        }
    }
}
