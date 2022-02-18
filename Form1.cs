using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GUICalendar
{
    public partial class Form1 : Form
    {
        static string[] Scopes = { CalendarService.Scope.Calendar,
                                   CalendarService.Scope.CalendarEvents };

        static string ApplicationName = System.AppDomain.CurrentDomain.FriendlyName;

        List<Account> Accounts = new List<Account>();
        List<Task> Tasks = new List<Task>();
        Events ev = null;
        bool[] doNotDisturb = new bool[96];
        //int[,] dayOfWeek = new int[5, 3];
        DateTime bestDay = DateTime.Now;
        bool taskDateSelected = false;
        bool taskDueDateSelected = false;
        int maxStart = 0;
        int maxCount = 0;

        public Form1()
        {
            InitializeComponent();
            goToBedBox.SelectedIndex = 20;
            wakeUpBox.SelectedIndex = 7;
            monthCalendar1.TodayDate = DateTime.Now;
            monthCalendar2.TodayDate = DateTime.Now;
            monthCalendar3.TodayDate = DateTime.Now;

            InitializeTimer();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();

            readExcelButton.PerformClick();

            form2.Close();
        }

        private void InitializeTimer()
        {
            timer1.Interval = 120000;
            timer1.Tick += new EventHandler(Timer1_Tick);

            timer1.Enabled = true;

            button1.Text = "Stop";
            button1.Click += new EventHandler(Button1_Click);
        }

        private void Timer1_Tick(object Sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();

            readExcelButton.PerformClick();

            form2.Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (button1.Text == "Stop")
            {
                button1.Text = "Start";
                timer1.Enabled = false;
            }
            else
            {
                button1.Text = "Stop";
                timer1.Enabled = true;
            }
        }


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        static extern System.IntPtr FindWindow(string lpClassName, string lpWindowName);

        //private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        //{
        //    this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected);
        //}


        static Events ListUpComing()
        {
            UserCredential credential;
            TimeSpan duration = new TimeSpan(30, 0, 0, 0);

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "admin",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            //request.TimeMax = DateTime.Now.Add(duration);
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 100;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            Events events = request.Execute();
            Console.WriteLine("Upcoming events:");
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    Console.WriteLine("{0} ({1})", eventItem.Summary, when);
                }
            }
            else
            {
                Console.WriteLine("No upcoming events found.");
            }
            Console.Read();

            return events;
        }

        //private void monthCalendar1_DateSelected(object sender, System.Windows.Forms.DateRangeEventArgs e)
        //{
        //    // Show the start date in the text box.
        //    this.textBox1.Text = "Date Selected: " + e.Start.ToShortDateString();
        //}
        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            // Show the start date in the text box.
            this.textBox1.Text = "Date Selected: " + e.Start.ToShortDateString();
        }

        private void monthCalendar2_DateSelected(object sender, DateRangeEventArgs e)
        {
            // Show the start date in the text box.
            this.textBox1.Text = "Date Selected: " + e.Start.ToShortDateString();
            taskDateSelected = true;
        }

        private void monthCalendar3_DateSelected(object sender, DateRangeEventArgs e)
        {
            this.textBox1.Text = "Date Selected: " + e.Start.ToShortDateString();
            taskDueDateSelected = true;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                DateTime dat = Convert.ToDateTime($"{ev.Items[listBox1.SelectedIndex].Start.DateTimeRaw}");
                richTextBox1.Text = $"Event: { ev.Items[listBox1.SelectedIndex].Summary} \n" +
                                    $"Start Time: { ev.Items[listBox1.SelectedIndex].Start.DateTime} \n" +
                                    $"End Time: { ev.Items[listBox1.SelectedIndex].End.DateTime} \n" +
                                    $"Date: {dat.ToString("dd")}";
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem != null)
            {
                richTextBox1.Text = $"Account: { Accounts[listBox2.SelectedIndex].Name} \n" +
                                    $"Due: { Accounts[listBox2.SelectedIndex].Due} \n" +
                                    $"Balance: { Accounts[listBox2.SelectedIndex].Balance} \n" +
                                    $"Minimum Payment: { Accounts[listBox2.SelectedIndex].MinPayment} \n" +
                                    $"DateTime Due: { Accounts[listBox2.SelectedIndex].DueDate.ToString("d") }";
            }
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox3.SelectedItem != null)
            {
                richTextBox1.Text = $"Task: { Tasks[listBox3.SelectedIndex].Name} \n" +
                                    $"Date: { Tasks[listBox3.SelectedIndex].Date} \n" +
                                    $"DateTime Date: { Tasks[listBox3.SelectedIndex].TaskDate.ToString("d") }";
            }
        }

        private void readExcelButton_Click(object sender, EventArgs e)
        {
            ev = ListUpComing();
            Accounts.Clear();
            Tasks.Clear();
            Accounts = AccountDB.ReadAccounts(Accounts);
            Tasks = AccountDB.ReadTasks(Tasks);
            AccountDB.AddEvents(ev);

            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            richTextBox1.Clear();

            for (int y = 0; y < ev.Items.Count; y++)
            {
                listBox1.Items.Add($"{ev.Items[y].Summary}");
            }
            for (int x = 0; x < Accounts.Count; x++)
            {
                listBox2.Items.Add($"{Accounts[x].Name}");
            }

            for (int y = 0; y < Tasks.Count; y++)
            {
                listBox3.Items.Add($"{Tasks[y].Name}");
            }
        }

        private void addAccountButton_Click(object sender, EventArgs e)
        {
            CreateNewAccount();
            accountBox.Clear();
            DateTime reset = DateTime.Today;
            monthCalendar1.SelectionStart = reset;
            monthCalendar1.SelectionEnd = reset;
            textBox1.Clear();
            balanceBox.Clear();
            minPaymentBox.Clear();
            richTextBox1.Clear();
        }

        public void CreateNewAccount()
        {
            Accounts.Clear();
            Accounts = AccountDB.ReadAccounts(Accounts);
            listBox2.Items.Clear();
            for (int x = 0; x < Accounts.Count; x++)
            {
                listBox2.Items.Add($"{Accounts[x].Name}");
            }

            Account a = new Account();

            a.Name = accountBox.Text;
            a.Due = Convert.ToInt32(monthCalendar1.SelectionRange.Start.ToString("dd"));
            a.DueDate = monthCalendar1.SelectionRange.Start;
            a.Balance = Convert.ToDouble(balanceBox.Text);
            a.MinPayment = Convert.ToDouble(minPaymentBox.Text);

            Accounts.Add(a);

            listBox2.Items.Add($"{a.Name}");

            a = AccountDB.Add(a, Accounts.Count);
        }

        private void addTaskButton_Click(object sender, EventArgs e)
        {
            CreateNewTask();
            taskBox.Clear();
            DateTime reset = DateTime.Today;
            monthCalendar2.SelectionStart = reset;
            monthCalendar2.SelectionEnd = reset;
            monthCalendar3.SelectionStart = reset;
            monthCalendar3.SelectionEnd = reset;
            taskDateSelected = false;
            taskDueDateSelected = false;
            bestDay = DateTime.Now;
            textBox1.Clear();
            richTextBox1.Clear();
        }

        public void CreateNewTask()
        {
            Tasks.Clear();
            Tasks = AccountDB.ReadTasks(Tasks);
            richTextBox3.Clear();
            listBox3.Items.Clear();
            for (int y = 0; y < Tasks.Count; y++)
            {
                listBox3.Items.Add($"{Tasks[y].Name}");
            }

            Task t = new Task();

            t.Name = taskBox.Text;

            if (!taskDateSelected)
            {

                button2.PerformClick();
                t.Date = Convert.ToInt32(bestDay.ToString("dd"));
                t.TaskDate = bestDay;
            }
            else
            {
                t.Date = Convert.ToInt32(monthCalendar2.SelectionRange.Start.ToString("dd"));
                t.TaskDate = monthCalendar2.SelectionRange.Start;
            }

            Tasks.Add(t);
            listBox3.Items.Add($"{t.Name}");

            t = AccountDB.AddTask(t, Tasks.Count);
        }


        private void excelFile_Click(object sender, EventArgs e)
        {
            string excelFile = "C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs";

            Excel.Application excelApp = new Excel.Application();

            try
            {
                Excel.Workbook wB = excelApp.Workbooks.Open(excelFile);
                Excel.Sheets allSheets = wB.Worksheets;

                Excel.Worksheet wS1 = (Worksheet)allSheets["Accounts"];
                Excel.Worksheet wS2 = (Worksheet)allSheets["Tasks"];
                Excel.Worksheet wS3 = (Worksheet)allSheets["Events"];

                wS1.Activate();
                Excel.Range wR = null;

                string caption = excelApp.Caption;
                IntPtr handler = FindWindow(null, caption);
                SetForegroundWindow(handler);
                excelApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("File Not Found");
            }
        }

        private void addAccountEventButton_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem != null)
            {
                UserCredential credential;

                using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                ev = ListUpComing();
                AccountDB.AddEvents(ev);
                listBox1.Items.Clear();
                for (int y = 0; y < ev.Items.Count; y++)
                {
                    listBox1.Items.Add($"{ev.Items[y].Summary}");
                }

                DateTime dat = Convert.ToDateTime(Accounts[listBox2.SelectedIndex].DueDate);
                DateTime datNow = DateTime.Now;

                if (DateTime.Compare(datNow, dat) < 0 || dat.ToString("dd") == datNow.ToString("dd"))
                {
                    DateTime gT = Convert.ToDateTime(GetTime(dat));
                    displayDateBox.Text = GetTime(dat);

                    Event newEvent = new Event()
                    {
                        Summary = $"{Accounts[listBox2.SelectedIndex].Name}",
                        Location = "Home",
                        Description = "Payment Information",
                        Start = new EventDateTime()
                        {
                            //DateTime = DateTime.Parse("2020-06-16T00:00:00-04:00"),

                            DateTime = DateTime.Parse($"{Convert.ToDateTime(Accounts[listBox2.SelectedIndex].DueDate).ToString("yyyy-MM-dd")}T" +
                                        $"{GetTime(Convert.ToDateTime(Accounts[listBox2.SelectedIndex].DueDate))}:00 -04:00"),

                            TimeZone = "America/Detroit",
                        },
                        End = new EventDateTime()
                        {
                            //DateTime = DateTime.Parse("2020-06-16T01:00:00-04:00"),

                            DateTime = DateTime.Parse($"{Convert.ToDateTime(Accounts[listBox2.SelectedIndex].DueDate).ToString("yyyy-MM-dd")}T" +
                                        $"{gT.AddMinutes(15).ToString("HH:mm")}:00 -04:00"),

                            TimeZone = "America/Detroit",
                        },

                        //Recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=2" },
                        //Attendees = new EventAttendee[] {
                        //new EventAttendee() { Email = "dpmjan21@gmail.com" },
                        //new EventAttendee() { Email = "sbrin@example.com" },
                        //},
                        Reminders = new Event.RemindersData()

                        {
                            UseDefault = false,
                            Overrides = new EventReminder[]
                        {
                    //new EventReminder() { Method = "email", Minutes = 24 * 60 },
                    //new EventReminder() { Method = "email", Minutes = 2 },
                    new EventReminder() { Method = "popup", Minutes = 2 },
                        }
                        }
                    };
                    String calendarId = "primary";
                    EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
                    Event createdEvent = request.Execute();
                    Console.WriteLine("Event created: {0}", createdEvent.HtmlLink);

                    ev = ListUpComing();
                    AccountDB.AddEvents(ev);
                    listBox1.Items.Clear();
                    for (int y = 0; y < ev.Items.Count; y++)
                    {
                        listBox1.Items.Add($"{ev.Items[y].Summary}");
                    }
                    richTextBox1.Clear();
                    listBox2.ClearSelected();
                }
                else
                {
                    MessageBox.Show($"Account: {Accounts[listBox2.SelectedIndex].Name}\nEvent not created\n{Accounts[listBox2.SelectedIndex].DueDate.ToString("d")}\nThis date is in the past");
                }
            }
            else
            {
                MessageBox.Show("No account selected");
            }
        }

        private void addTaskEventButton_Click(object sender, EventArgs e)
        {
            if (listBox3.SelectedItem != null)
            {
                UserCredential credential;

                using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                ev = ListUpComing();
                AccountDB.AddEvents(ev);
                listBox1.Items.Clear();
                for (int y = 0; y < ev.Items.Count; y++)
                {
                    listBox1.Items.Add($"{ev.Items[y].Summary}");
                }

                DateTime dat = Convert.ToDateTime(Tasks[listBox3.SelectedIndex].TaskDate);
                DateTime datNow = DateTime.Now;

                if (DateTime.Compare(datNow, dat) < 0 || dat.ToString("dd") == datNow.ToString("dd"))
                {
                    if (GetTime(dat) == "none")
                    {
                        MessageBox.Show($"No Available Time");
                    }
                    else
                    {


                        DateTime gT = Convert.ToDateTime(GetTime(dat));
                        //displayDateBox.Text = GetTime(dat);
                        displayDateBox.Text = gT.ToString("dd");

                        Event newEvent = new Event()
                        {
                            Summary = $"{Tasks[listBox3.SelectedIndex].Name}",
                            Location = "Home",
                            Description = "Payment Information",
                            Start = new EventDateTime()
                            {
                                //DateTime = DateTime.Parse("2020-06-16T00:00:00-04:00"),

                                DateTime = DateTime.Parse($"{Convert.ToDateTime(Tasks[listBox3.SelectedIndex].TaskDate).ToString("yyyy-MM-dd")}T" +
                                            $"{GetTime(Convert.ToDateTime(Tasks[listBox3.SelectedIndex].TaskDate))}:00 -04:00"),

                                TimeZone = "America/Detroit",
                            },
                            End = new EventDateTime()
                            {
                                //DateTime = DateTime.Parse("2020-06-16T01:00:00-04:00"),

                                DateTime = DateTime.Parse($"{Convert.ToDateTime(Tasks[listBox3.SelectedIndex].TaskDate).ToString("yyyy-MM-dd")}T" +
                                            $"{gT.AddMinutes(15).ToString("HH:mm")}:00 -04:00"),

                                TimeZone = "America/Detroit",
                            },

                            //Recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=2" },
                            //Attendees = new EventAttendee[] {
                            //new EventAttendee() { Email = "dpmjan21@gmail.com" },
                            //new EventAttendee() { Email = "sbrin@example.com" },
                            //},
                            Reminders = new Event.RemindersData()

                            {
                                UseDefault = false,
                                Overrides = new EventReminder[]
                            {
                    //new EventReminder() { Method = "email", Minutes = 24 * 60 },
                    //new EventReminder() { Method = "email", Minutes = 2 },
                    new EventReminder() { Method = "popup", Minutes = 2 },
                            }
                            }
                        };
                        String calendarId = "primary";
                        EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
                        Event createdEvent = request.Execute();
                        Console.WriteLine("Event created: {0}", createdEvent.HtmlLink);

                        ev = ListUpComing();
                        AccountDB.AddEvents(ev);
                        listBox1.Items.Clear();
                        for (int y = 0; y < ev.Items.Count; y++)
                        {
                            listBox1.Items.Add($"{ev.Items[y].Summary}");
                        }
                        richTextBox1.Clear();
                        listBox3.ClearSelected();
                    }
                }
                else
                {
                    MessageBox.Show($"Task: {Tasks[listBox3.SelectedIndex].Name}\nEvent not created\n{Tasks[listBox3.SelectedIndex].TaskDate.ToString("d")}\nThis date is in the past");
                }
            }
            else
            {
                MessageBox.Show("No task selected");
            }
        }

        private void addTasksButton_Click(object sender, EventArgs e)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            for (int z = 0; z < Accounts.Count; z++)
            {
                ev = ListUpComing();
                AccountDB.AddEvents(ev);
                listBox1.Items.Clear();
                for (int y = 0; y < ev.Items.Count; y++)
                {
                    listBox1.Items.Add($"{ev.Items[y].Summary}");
                }

                DateTime dat = Convert.ToDateTime(Accounts[z].DueDate);
                DateTime datNow = DateTime.Now;

                if (DateTime.Compare(datNow, dat) < 0 || dat.ToString("dd") == datNow.ToString("dd"))
                {

                    //DateTime gTa = Convert.ToDateTime(GetTime(Accounts[z].DueDate));
                    DateTime gTa = Convert.ToDateTime(GetTime(dat));

                    Event newEvent = new Event()
                    {
                        Summary = $"{Accounts[z].Name}",
                        Location = "Home",
                        Description = "Payment Information",
                        Start = new EventDateTime()
                        {
                            //DateTime = DateTime.Parse("2020-06-16T00:00:00-04:00"),

                            DateTime = DateTime.Parse($"{Convert.ToDateTime(Accounts[z].DueDate).ToString("yyyy-MM-dd")}T" +
                                    $"{GetTime(Convert.ToDateTime(Accounts[z].DueDate))}:00 -04:00"),

                            TimeZone = "America/Detroit",
                        },
                        End = new EventDateTime()
                        {
                            //DateTime = DateTime.Parse("2020-06-16T01:00:00-04:00"),

                            DateTime = DateTime.Parse($"{Convert.ToDateTime(Accounts[z].DueDate).ToString("yyyy-MM-dd")}T" +
                                    $"{gTa.AddMinutes(15).ToString("HH:mm")}:00 -04:00"),

                            TimeZone = "America/Detroit",
                        },

                        //Recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=2" },
                        //Attendees = new EventAttendee[] {
                        //new EventAttendee() { Email = "dpmjan21@gmail.com" },
                        //new EventAttendee() { Email = "sbrin@example.com" },
                        //},
                        Reminders = new Event.RemindersData()
                        {
                            UseDefault = false,
                            Overrides = new EventReminder[]
                        {
                    //new EventReminder() { Method = "email", Minutes = 24 * 60 },
                    //new EventReminder() { Method = "sms", Minutes = 10 },
                    new EventReminder() { Method = "popup", Minutes = 2 },
                        }
                        }
                    };
                    String calendarId = "primary";
                    EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
                    Event createdEvent = request.Execute();
                    Console.WriteLine("Event created: {0}", createdEvent.HtmlLink);
                }
                else
                {
                    MessageBox.Show($"Account: {Accounts[z].Name}\nEvent not created\n{Accounts[z].DueDate.ToString("d")}\nThis date is in the past");
                }
            }

            for (int z = 0; z < Tasks.Count; z++)
            {
                ev = ListUpComing();
                AccountDB.AddEvents(ev);
                listBox1.Items.Clear();
                for (int y = 0; y < ev.Items.Count; y++)
                {
                    listBox1.Items.Add($"{ev.Items[y].Summary}");
                }

                DateTime dat = Convert.ToDateTime(Tasks[z].TaskDate);
                DateTime datNow = DateTime.Now;

                if (DateTime.Compare(datNow, dat) < 0 || dat.ToString("dd") == datNow.ToString("dd"))
                {

                    //DateTime gTt = Convert.ToDateTime(GetTime(Tasks[z].TaskDate));
                    DateTime gTt = Convert.ToDateTime(GetTime(dat));

                    Event newEvent = new Event()
                    {
                        Summary = $"{Tasks[z].Name}",
                        Location = "Home",
                        Description = "Payment Information",
                        Start = new EventDateTime()
                        {
                            //DateTime = DateTime.Parse("2020-06-16T00:00:00-04:00"),

                            DateTime = DateTime.Parse($"{Convert.ToDateTime(Tasks[z].TaskDate).ToString("yyyy-MM-dd")}T" +
                                    $"{GetTime(Convert.ToDateTime(Tasks[z].TaskDate))}:00 -04:00"),

                            TimeZone = "America/Detroit",
                        },
                        End = new EventDateTime()
                        {
                            //DateTime = DateTime.Parse("2020-06-16T01:00:00-04:00"),

                            DateTime = DateTime.Parse($"{Convert.ToDateTime(Tasks[z].TaskDate).ToString("yyyy-MM-dd")}T" +
                                    $"{gTt.AddMinutes(15).ToString("HH:mm")}:00 -04:00"),

                            TimeZone = "America/Detroit",
                        },

                        //Recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=2" },
                        //Attendees = new EventAttendee[] {
                        //new EventAttendee() { Email = "dpmjan21@gmail.com" },
                        //new EventAttendee() { Email = "sbrin@example.com" },
                        //},
                        Reminders = new Event.RemindersData()
                        {
                            UseDefault = false,
                            Overrides = new EventReminder[]
                        {
                    //new EventReminder() { Method = "email", Minutes = 24 * 60 },
                    //new EventReminder() { Method = "sms", Minutes = 10 },
                    new EventReminder() { Method = "popup", Minutes = 2 },
                        }
                        }
                    };
                    String calendarId = "primary";
                    EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
                    Event createdEvent = request.Execute();
                    Console.WriteLine("Event created: {0}", createdEvent.HtmlLink);
                }
                else
                {
                    MessageBox.Show($"Task: {Tasks[z].Name}\nEvent not created\n{Tasks[z].TaskDate.ToString("d")}\nThis date is in the past");
                }

                ev = ListUpComing();
                AccountDB.AddEvents(ev);
                listBox1.Items.Clear();
                for (int y = 0; y < ev.Items.Count; y++)
                {
                    listBox1.Items.Add($"{ev.Items[y].Summary}");
                }
                richTextBox1.Clear();
            }
        }

        private void doNotDisturbButton_Click(object sender, EventArgs e)
        {
            //DateTime dat = Convert.ToDateTime(Accounts[listBox2.SelectedIndex].DueDate);
            DateTime dat = Convert.ToDateTime(Tasks[listBox3.SelectedIndex].TaskDate);
            //displayDateBox.Text = Convert.ToString(dat);
            displayDateBox.Text = GetTime(dat);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int spanOfDays;
            if(!taskDueDateSelected)
            {
                spanOfDays = 3;
            }
            else
            {
                spanOfDays = Convert.ToInt32((monthCalendar3.SelectionRange.Start - DateTime.Today).TotalDays + 1);
            }

            int[,] dayOfWeek = new int[spanOfDays, 3];
            Array.Clear(dayOfWeek, 0, dayOfWeek.Length);
            DateTime[] bDate = new DateTime[spanOfDays];
            int mCount = 0;
            int dayIndex = 0;
            DateTime dayOne = DateTime.Now;
            DateTime nextDays = DateTime.Today;

            for (int d = 0; d < spanOfDays; d++)
            {
                if (d == 0)
                {
                    bDate[d] = dayOne;
                    GetTime(dayOne);
                    dayOfWeek[d, 0] = Convert.ToInt32(dayOne.ToString("dd"));
                    //dayOfWeek[d, 1] = Convert.ToInt32(maxStartBox.Text);
                    //dayOfWeek[d, 2] = Convert.ToInt32(maxCountBox.Text);
                    dayOfWeek[d, 1] = maxStart;
                    dayOfWeek[d, 2] = maxCount;
                }
                else
                {
                    nextDays = nextDays.AddDays(1);
                    bDate[d] = nextDays;
                    GetTime(nextDays);
                    dayOfWeek[d, 0] = Convert.ToInt32(nextDays.ToString("dd"));
                    //dayOfWeek[d, 1] = Convert.ToInt32(maxStartBox.Text);
                    //dayOfWeek[d, 2] = Convert.ToInt32(maxCountBox.Text);
                    dayOfWeek[d, 1] = maxStart;
                    dayOfWeek[d, 2] = maxCount;
                }
            }

            for (int x = 0; x < spanOfDays; x++)
            {

                if (x == 0)
                {
                    richTextBox3.Text = $"Date: {dayOfWeek[x, 0]}   Start: {dayOfWeek[x, 1]}   Count: {dayOfWeek[x, 2]}";
                }
                else
                {
                    richTextBox3.AppendText($"\nDate: {dayOfWeek[x, 0]}   Start: {dayOfWeek[x, 1]}   Count: {dayOfWeek[x, 2]}");
                }
            }

            for (int w = 0; w < spanOfDays; w++)
            {
                if (dayOfWeek[w, 2] > mCount)
                {
                    mCount = dayOfWeek[w, 2];
                    dayIndex = w;
                }
            }

            richTextBox3.AppendText($"\n\nBest Day And Time\nDate: {dayOfWeek[dayIndex, 0]}   Start: {dayOfWeek[dayIndex, 1]}   Count: {dayOfWeek[dayIndex, 2]}");

            bestDay = bDate[dayIndex];
        }

        public string GetTime(DateTime dateInquire)
        {
            string tH = "00";
            string tM = "00";
            bool timeSelected = false;
            string time = null;
            bool procede = false;
            int freeCount = 0;
            int freeStart = 0;
            maxStart = 0;
            maxCount = 0;
            int s = goToBedBox.SelectedIndex;
            int w = wakeUpBox.SelectedIndex;
            int eSh = 0;
            int eEh = 0;
            int eSm = 0;
            int eEm = 0;
            string dndDate = dateInquire.ToString("dd");
            string dndNowDate = DateTime.Now.ToString("dd");
            int dndNowHours = Convert.ToInt32(DateTime.Now.ToString("HH"));
            int dndNowMinutes = Convert.ToInt32(DateTime.Now.ToString("mm"));


            // Prevents past time on current day from being chosen
            if (dndDate == dndNowDate)
            {
                for (int d = 0; d < dndNowHours; d++)
                {
                    doNotDisturb[d * 4] = true;
                    doNotDisturb[(d * 4) + 1] = true;
                    doNotDisturb[(d * 4) + 2] = true;
                    doNotDisturb[(d * 4) + 3] = true;
                }

                if (dndNowMinutes > 44)
                {
                    doNotDisturb[dndNowHours * 4] = true;
                    doNotDisturb[(dndNowHours * 4) + 1] = true;
                    doNotDisturb[(dndNowHours * 4) + 2] = true;
                    doNotDisturb[(dndNowHours * 4) + 3] = true;
                }
                else if (dndNowMinutes > 29)
                {
                    doNotDisturb[dndNowHours * 4] = true;
                    doNotDisturb[(dndNowHours * 4) + 1] = true;
                    doNotDisturb[(dndNowHours * 4) + 2] = true;
                }
                else if (dndNowMinutes > 14)
                {
                    doNotDisturb[dndNowHours * 4] = true;
                    doNotDisturb[(dndNowHours * 4) + 1] = true;
                }
                else
                {
                    doNotDisturb[dndNowHours * 4] = true;
                }
            }

            // Prevents time during do not disturb hours from being chosen
            if (s < w)
            {
                for (int ss = s; ss < w; ss++)
                {
                    doNotDisturb[ss * 4] = true;
                    doNotDisturb[(ss * 4) + 1] = true;
                    doNotDisturb[(ss * 4) + 2] = true;
                    doNotDisturb[(ss * 4) + 3] = true;
                }
            }
            else
            {
                for (int ss = 0; ss < w; ss++)
                {
                    doNotDisturb[ss * 4] = true;
                    doNotDisturb[(ss * 4) + 1] = true;
                    doNotDisturb[(ss * 4) + 2] = true;
                    doNotDisturb[(ss * 4) + 3] = true;
                }
                for (int ss = s; ss < 24; ss++)
                {
                    doNotDisturb[ss * 4] = true;
                    doNotDisturb[(ss * 4) + 1] = true;
                    doNotDisturb[(ss * 4) + 2] = true;
                    doNotDisturb[(ss * 4) + 3] = true;
                }
            }

            // Prevents time during scheduled events from being chosen
            for (int b = 0; b < ev.Items.Count; b++)
            {
                DateTime da = Convert.ToDateTime($"{ev.Items[b].Start.DateTimeRaw}");

                if (da.ToString("dd") == dndDate)
                {
                    DateTime dteS = DateTime.Parse($"{ev.Items[b].Start.DateTimeRaw}");
                    DateTime dteE = DateTime.Parse($"{ev.Items[b].End.DateTimeRaw}");
                    eSh = Convert.ToInt32(dteS.ToString("HH"));
                    eSm = Convert.ToInt32(dteS.ToString("mm"));
                    eEh = Convert.ToInt32(dteE.ToString("HH"));
                    eEm = Convert.ToInt32(dteE.ToString("mm"));

                    if (eSh <= eEh)
                    {
                        for (int ss = eSh; ss <= eEh; ss++)
                        {
                            if (eSh == eEh)
                            {
                                if (eSm == 0)
                                {
                                    if (eEm > 45)
                                    {
                                        doNotDisturb[eEh * 4] = true;
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                        doNotDisturb[(eEh * 4) + 2] = true;
                                        doNotDisturb[(eEh * 4) + 3] = true;
                                    }
                                    else if (eEm > 30)
                                    {
                                        doNotDisturb[eEh * 4] = true;
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                        doNotDisturb[(eEh * 4) + 2] = true;
                                    }
                                    else if (eEm > 15)
                                    {
                                        doNotDisturb[eEh * 4] = true;
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                    }
                                    else if (eEm > 0)
                                    {
                                        doNotDisturb[eEh * 4] = true;
                                    }
                                }
                                else if (eSm == 15)
                                {
                                    if (eEm > 45)
                                    {
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                        doNotDisturb[(eEh * 4) + 2] = true;
                                        doNotDisturb[(eEh * 4) + 3] = true;
                                    }
                                    else if (eEm > 30)
                                    {
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                        doNotDisturb[(eEh * 4) + 2] = true;
                                    }
                                    else if (eEm > 15)
                                    {
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                    }
                                }
                                else if (eSm == 30)
                                {
                                    if (eEm > 45)
                                    {
                                        doNotDisturb[(eEh * 4) + 2] = true;
                                        doNotDisturb[(eEh * 4) + 3] = true;
                                    }
                                    else if (eEm > 30)
                                    {
                                        doNotDisturb[(eEh * 4) + 2] = true;
                                    }
                                }
                                else if (eSm == 45)
                                {
                                    if (eEm > 45)
                                    {
                                        doNotDisturb[(eEh * 4) + 3] = true;
                                    }
                                }
                            }
                            else
                            {
                                if (ss == eSh)
                                {
                                    if (eSm >= 45)
                                    {
                                        doNotDisturb[(eSh * 4) + 3] = true;
                                    }
                                    else if (eSm >= 30)
                                    {
                                        doNotDisturb[(eSh * 4) + 2] = true;
                                        doNotDisturb[(eSh * 4) + 3] = true;
                                    }
                                    else if (eSm >= 15)
                                    {
                                        doNotDisturb[(eSh * 4) + 1] = true;
                                        doNotDisturb[(eSh * 4) + 2] = true;
                                        doNotDisturb[(eSh * 4) + 3] = true;
                                    }
                                    else if (eSm >= 0)
                                    {
                                        doNotDisturb[eSh * 4] = true;
                                        doNotDisturb[(eSh * 4) + 1] = true;
                                        doNotDisturb[(eSh * 4) + 2] = true;
                                        doNotDisturb[(eSh * 4) + 3] = true;
                                    }
                                }
                                else if (ss == eEh)
                                {
                                    if (eEm > 45)
                                    {
                                        doNotDisturb[(eEh * 4)] = true;
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                        doNotDisturb[(eEh * 4) + 2] = true;
                                        doNotDisturb[(eEh * 4) + 3] = true;
                                    }
                                    else if (eEm > 30)
                                    {
                                        doNotDisturb[(eEh * 4)] = true;
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                        doNotDisturb[(eEh * 4) + 2] = true;
                                    }
                                    else if (eEm > 15)
                                    {
                                        doNotDisturb[(eEh * 4)] = true;
                                        doNotDisturb[(eEh * 4) + 1] = true;
                                    }
                                    else if (eEm > 0)
                                    {
                                        doNotDisturb[(eEh * 4)] = true;
                                    }
                                }
                                else
                                {
                                    doNotDisturb[ss * 4] = true;
                                    doNotDisturb[(ss * 4) + 1] = true;
                                    doNotDisturb[(ss * 4) + 2] = true;
                                    doNotDisturb[(ss * 4) + 3] = true;
                                }
                            }
                        }
                    }
                }
            }

            for (int x = 0; x < 96; x++)
            {
                if (x == 0)
                {
                    richTextBox2.Text = $"Index {x}: {doNotDisturb[x].ToString()} ";
                }
                else
                {
                    richTextBox2.AppendText($"\nIndex {x}: {doNotDisturb[x].ToString()}");
                }
            }

            for (int x = 0; x < 96; x++)
            {
                if (!doNotDisturb[x])
                {
                    freeStart = x;
                    procede = true;

                    for (int y = freeStart; y < 96; y++)
                    {
                        if (!doNotDisturb[y] && procede)
                        {
                            freeCount += 1;
                        }
                        else
                        {
                            procede = false;
                        }
                    }

                    if (freeCount > maxCount)
                    {
                        maxCount = freeCount;
                        maxStart = freeStart;
                        timeSelected = true;
                    }
                    freeCount = 0;
                }
            }

            for (int dd = 0; dd < 96; dd++)
            {
                doNotDisturb[dd] = false;
            }

            if (!timeSelected)
            {
                //MessageBox.Show($"No Available Time");
                return "none";
            }
            else
            {
                timeSelected = false;



                maxStartBox.Text = $"Max Start: {maxStart}";
                maxCountBox.Text = $"Max Count: {maxCount}";

                if (maxStart / 4 < 10)
                {
                    tH = "0" + Convert.ToString(maxStart / 4);
                }
                else
                {
                    tH = Convert.ToString(maxStart / 4);
                }

                if (maxStart % 4 == 0)
                {
                    tM = "00";
                }
                else if (maxStart % 4 == 1)
                {
                    tM = "15";
                }
                else if (maxStart % 4 == 2)
                {
                    tM = "30";
                }
                else if (maxStart % 4 == 3)
                {
                    tM = "45";
                }

                time = tH + ":" + tM;

                return time;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DateTime reset = DateTime.Today;
            monthCalendar2.SelectionStart = reset;
            monthCalendar2.SelectionEnd = reset;
            monthCalendar3.SelectionStart = reset;
            monthCalendar3.SelectionEnd = reset;
            taskDateSelected = false;
            taskDueDateSelected = false;
            richTextBox2.Clear();
            richTextBox3.Clear();
        }
    }
}
