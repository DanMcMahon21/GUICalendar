using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GUICalendar
{
    public class TaskN
    {
        private string t_name;
        private int t_date;
        private DateTime t_taskDate;

        public TaskN()
        {
            t_name = null;
            t_date = 1;
            t_taskDate = DateTime.Now;
        }

        public string Name
        {
            get
            {
                return t_name;
            }
            set
            {
                value = value.ToUpper();
                t_name = value;
            }
        }

        public int Date
        {
            get
            {
                return t_date;
            }
            set
            {
                if (value >= 0)
                    t_date = value;
            }
        }

        public DateTime TaskDate
        {
            get
            {
                return t_taskDate;
            }
            set
            {
                t_taskDate = value;
            }
        }

        public override string ToString()
        {
            return $"Task: {t_name}\nDate: {t_date}";
        }
    }
}
