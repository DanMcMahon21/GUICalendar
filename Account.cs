using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GUICalendar
{
    public class Account
    {
        private string a_name;
        private double a_balance;
        private double a_minPayment;
        private int a_due;
        private DateTime a_dueDate;

        public Account()
        {
            a_name = null;
            a_balance = 0;
            a_minPayment = 0;
            a_due = 0;
            a_dueDate = DateTime.Now;
        }

        public string Name
        {
            get
            {
                return a_name;
            }
            set
            {
                value = value.ToUpper();
                a_name = value;
            }
        }

        public double Balance
        {
            get
            {
                return a_balance;
            }
            set
            {
                if (value >= 0)
                    a_balance = value;
            }
        }

        public double MinPayment
        {
            get
            {
                return a_minPayment;
            }
            set
            {
                if (value >= 0)
                    a_minPayment = value;
            }
        }

        public int Due
        {
            get
            {
                return a_due;
            }
            set
            {
                if (value >= 0)
                    a_due = value;
            }
        }

        public DateTime DueDate
        {
            get
            {
                return a_dueDate;
            }
            set
            {
                a_dueDate = value;
            }
        }

        public override string ToString()
        {
            return $"Account: {a_name}\nBalance: {a_balance}\nMinimum Payment: {a_minPayment}\nDue: {a_due}";
        }
    }
}
