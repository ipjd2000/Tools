using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LLDOCX.Helpers
{
    public class DocumentHeadingData
    {
        public DocumentHeadingData()
        {
            SubHeadings = new List<string>();
        }
        public int Index { get; set; }
        public string Heading { get; set; }
        public bool IsRequired { get; set; }
        public List<string> SubHeadings { get; set; }
    }
}
