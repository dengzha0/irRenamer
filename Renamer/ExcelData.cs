using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Renamer
{
    public class ExcelData
    {
        public int type;
        public string oldName;
        public string oldExt;
        public string newName;
        public string newExt;

        public ExcelData()
        {
            type = 0;
            oldName = string.Empty;
            oldExt = string.Empty;
            newName = string.Empty;
            newExt = string.Empty;
        }
    }
}
