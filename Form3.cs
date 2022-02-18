using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GUICalendar
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();

            for (int x = 0; x < Form1.Accounts.Count; x++)
            {
                listBox1.Items.Add($"{Form1.Accounts[x].Name}     " +
                    $"Due: {Form1.Accounts[x].DueDate.ToString("d")}     " +
                    $"Balance: {Form1.Accounts[x].Balance.ToString("C")}     " +
                    $"Minimum Payment: {Form1.Accounts[x].MinPayment.ToString("C")}");
            }
            for (int y = 0; y < Form1.Tasks.Count; y++)
            {
                listBox2.Items.Add($"{Form1.Tasks[y].Name}     " +
                    $"Date: {Form1.Tasks[y].TaskDate.ToString("d")}");
            }
            if (Form1.ev != null)
            {


                if (Form1.ev.Items == null || Form1.ev.Items.Count <= 0)
                {
                    MessageBox.Show($"No upcoming events found");
                }
                else
                {
                    AccountDB.AddEvents(Form1.ev);
                    for (int y = 0; y < Form1.ev.Items.Count; y++)
                    {
                        listBox3.Items.Add($"{Form1.ev.Items[y].Summary}     " +
                            $"Date: {Convert.ToDateTime(Form1.ev.Items[y].Start.DateTime).ToString("d")}     " +
                            $"Start Time: {Convert.ToDateTime(Form1.ev.Items[y].Start.DateTime).ToString("h:mm tt")}     " +
                            $"End Time: { Convert.ToDateTime(Form1.ev.Items[y].End.DateTime).ToString("h:mm tt")}");
                    }
                }
            }
            else
            {
                MessageBox.Show("No Credentials Entered");
            }
        }
    }
}
