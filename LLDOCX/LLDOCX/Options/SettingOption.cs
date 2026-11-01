using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LLDOCX.Options
{
    [Serializable]
   public class SettingOption
    {
        public bool CheckPageSize { get; set; } = true;
        public bool CheckPageOrientation { get; set; } = true;
        public bool CheckMarkup { get; set; } = true;
        public bool CheckFont { get; set; } = true;
        public bool CheckHyperlink { get; set; } = true;
        public bool CheckComment { get; set; } = true;
        public bool CheckBookmarker { get; set; } = true;
        public bool RemoveDocumentProperty { get; set; } = true;
        public bool KeepTrackingOn { get; set; } = true;
        public bool HeadingValidation { get; set; } = true;
        public bool CheckedAll { get; set; } = true;
        public string DefaultFontName { get; set; } = "Times New Roman";
    }
}
