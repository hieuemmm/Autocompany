using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCompany
{
    class App
    {
        public static int TabCurrent = (int)TABCONTROL.TAB_INFO;
        public static int SelectListBoxCurrent = (int)LISTBOX.NONE;

        public static int ToTalCompany = 0;
        public static int ToTalLink = 0;
        public static int TookLink = 0;
        public static int ExcelCompany = 0;

        public enum TABCONTROL
        {
            TAB_INFO,
            TAB_TEMPLATE,
            TAB_INFO_APP,
        }
        public enum LINKPAGE
        {
            LIST,
            SELECTED
        }
        public enum LISTBOX
        {
            NONE,
            LIST,
            SELECTED
        }
    }
}
