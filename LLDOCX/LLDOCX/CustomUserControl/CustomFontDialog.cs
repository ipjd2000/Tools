using LLDOCX.Options;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LLDOCX.CustomUserControl
{
    public partial class CustomFontDialog : Form
    {
        public string SelectedFont { get; set; }
        private List<string> supportedFonts;
        public CustomFontDialog(SettingOption settingOption)
        {
            InitializeComponent();
            supportedFonts = new List<string>() {"Arial","Arial Black","Arial Narrow","Arial Rounded Mt","Arial Unicode Ms","Batang","Batangche",
"Calibri","Cambria","Cambria Math","Castellar","Courier New","Georgia","Liberation Sans","Lucida Sans","MingLiU ExtB","Monaco","MS mincho",
"MS pmincho","PmingLiU ExtB","STIXGeneral","Symbol","Tahoma","Times","Times New Roman","Trebuchet MS","Verdana","Wide Latin", };
            cmbFonts.DataSource = supportedFonts;
            int index = supportedFonts.IndexOf(settingOption.DefaultFontName);
            cmbFonts.SelectedIndex = index;
        }

        private void OnbtnOKClicked(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            SelectedFont = cmbFonts.Text;
        }
    }
}
