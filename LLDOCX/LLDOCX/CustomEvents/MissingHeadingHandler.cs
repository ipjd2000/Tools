using LLDOCX.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LLDOCX.CustomEvents
{
    public class MissingHeadingHandler
    {

        public event EventHandler<List<DocumentHeadingData>> OnUpdateMissingHeadingData;
        public MissingHeadingHandler()
        {
            // OnEmailItemChanged("");
        }
        public void OnEmailItemChanged(List<DocumentHeadingData> e)
        {
            EventHandler<List<DocumentHeadingData>> handler = OnUpdateMissingHeadingData;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
