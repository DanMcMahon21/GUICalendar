namespace GUICalendar
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.addAccountButton = new System.Windows.Forms.Button();
            this.accountBox = new System.Windows.Forms.TextBox();
            this.displayDateBox = new System.Windows.Forms.TextBox();
            this.balanceBox = new System.Windows.Forms.TextBox();
            this.accountNameLabel = new System.Windows.Forms.Label();
            this.accountDueLabel = new System.Windows.Forms.Label();
            this.accountBalanceLabel = new System.Windows.Forms.Label();
            this.minPaymentBox = new System.Windows.Forms.TextBox();
            this.accountMinPaymentLabel = new System.Windows.Forms.Label();
            this.excelFile = new System.Windows.Forms.Button();
            this.readExcelButton = new System.Windows.Forms.Button();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.addAccountEventButton = new System.Windows.Forms.Button();
            this.listBox3 = new System.Windows.Forms.ListBox();
            this.addTasksButton = new System.Windows.Forms.Button();
            this.goToBedBox = new System.Windows.Forms.ComboBox();
            this.wakeUpBox = new System.Windows.Forms.ComboBox();
            this.doNotDisturbButton = new System.Windows.Forms.Button();
            this.taskBox = new System.Windows.Forms.TextBox();
            this.monthCalendar2 = new System.Windows.Forms.MonthCalendar();
            this.addTaskButton = new System.Windows.Forms.Button();
            this.accountListLabel = new System.Windows.Forms.Label();
            this.taskListLabel = new System.Windows.Forms.Label();
            this.taskNameLabel = new System.Windows.Forms.Label();
            this.taskDateLabel = new System.Windows.Forms.Label();
            this.eventListLabel = new System.Windows.Forms.Label();
            this.descriptionLabel = new System.Windows.Forms.Label();
            this.goToSleepLabel = new System.Windows.Forms.Label();
            this.wakeUpTimeLabel = new System.Windows.Forms.Label();
            this.addTaskEventButton = new System.Windows.Forms.Button();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.maxStartBox = new System.Windows.Forms.TextBox();
            this.maxCountBox = new System.Windows.Forms.TextBox();
            this.richTextBox3 = new System.Windows.Forms.RichTextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.monthCalendar3 = new System.Windows.Forms.MonthCalendar();
            this.button3 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.LightCyan;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(407, 1171);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(311, 35);
            this.textBox1.TabIndex = 0;
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.BackColor = System.Drawing.Color.LightCyan;
            this.monthCalendar1.Location = new System.Drawing.Point(15, 637);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 1;
            this.monthCalendar1.TodayDate = new System.DateTime(((long)(0)));
            this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected);
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.Color.LightCyan;
            this.richTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox1.Location = new System.Drawing.Point(1289, 97);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(356, 281);
            this.richTextBox1.TabIndex = 2;
            this.richTextBox1.Text = "";
            // 
            // listBox1
            // 
            this.listBox1.BackColor = System.Drawing.Color.LightCyan;
            this.listBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 29;
            this.listBox1.Location = new System.Drawing.Point(799, 97);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(419, 787);
            this.listBox1.TabIndex = 5;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // addAccountButton
            // 
            this.addAccountButton.BackColor = System.Drawing.Color.Teal;
            this.addAccountButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.addAccountButton.Location = new System.Drawing.Point(15, 1072);
            this.addAccountButton.Name = "addAccountButton";
            this.addAccountButton.Size = new System.Drawing.Size(311, 108);
            this.addAccountButton.TabIndex = 7;
            this.addAccountButton.Text = "Add Account";
            this.addAccountButton.UseVisualStyleBackColor = false;
            this.addAccountButton.Click += new System.EventHandler(this.addAccountButton_Click);
            // 
            // accountBox
            // 
            this.accountBox.BackColor = System.Drawing.Color.LightCyan;
            this.accountBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.accountBox.Location = new System.Drawing.Point(15, 557);
            this.accountBox.Name = "accountBox";
            this.accountBox.Size = new System.Drawing.Size(311, 35);
            this.accountBox.TabIndex = 8;
            // 
            // displayDateBox
            // 
            this.displayDateBox.BackColor = System.Drawing.Color.LightCyan;
            this.displayDateBox.Location = new System.Drawing.Point(1289, 820);
            this.displayDateBox.Name = "displayDateBox";
            this.displayDateBox.Size = new System.Drawing.Size(356, 26);
            this.displayDateBox.TabIndex = 10;
            // 
            // balanceBox
            // 
            this.balanceBox.BackColor = System.Drawing.Color.LightCyan;
            this.balanceBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.balanceBox.Location = new System.Drawing.Point(15, 937);
            this.balanceBox.Name = "balanceBox";
            this.balanceBox.Size = new System.Drawing.Size(311, 35);
            this.balanceBox.TabIndex = 11;
            // 
            // accountNameLabel
            // 
            this.accountNameLabel.AutoSize = true;
            this.accountNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.accountNameLabel.Location = new System.Drawing.Point(9, 517);
            this.accountNameLabel.Name = "accountNameLabel";
            this.accountNameLabel.Size = new System.Drawing.Size(210, 36);
            this.accountNameLabel.TabIndex = 12;
            this.accountNameLabel.Text = "Account Name";
            // 
            // accountDueLabel
            // 
            this.accountDueLabel.AutoSize = true;
            this.accountDueLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.accountDueLabel.Location = new System.Drawing.Point(9, 597);
            this.accountDueLabel.Name = "accountDueLabel";
            this.accountDueLabel.Size = new System.Drawing.Size(69, 36);
            this.accountDueLabel.TabIndex = 13;
            this.accountDueLabel.Text = "Due";
            // 
            // accountBalanceLabel
            // 
            this.accountBalanceLabel.AutoSize = true;
            this.accountBalanceLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.accountBalanceLabel.Location = new System.Drawing.Point(9, 897);
            this.accountBalanceLabel.Name = "accountBalanceLabel";
            this.accountBalanceLabel.Size = new System.Drawing.Size(122, 36);
            this.accountBalanceLabel.TabIndex = 14;
            this.accountBalanceLabel.Text = "Balance";
            // 
            // minPaymentBox
            // 
            this.minPaymentBox.BackColor = System.Drawing.Color.LightCyan;
            this.minPaymentBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.minPaymentBox.Location = new System.Drawing.Point(15, 1017);
            this.minPaymentBox.Name = "minPaymentBox";
            this.minPaymentBox.Size = new System.Drawing.Size(311, 35);
            this.minPaymentBox.TabIndex = 15;
            // 
            // accountMinPaymentLabel
            // 
            this.accountMinPaymentLabel.AutoSize = true;
            this.accountMinPaymentLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.accountMinPaymentLabel.Location = new System.Drawing.Point(9, 977);
            this.accountMinPaymentLabel.Name = "accountMinPaymentLabel";
            this.accountMinPaymentLabel.Size = new System.Drawing.Size(257, 36);
            this.accountMinPaymentLabel.TabIndex = 16;
            this.accountMinPaymentLabel.Text = "Minimum Payment";
            // 
            // excelFile
            // 
            this.excelFile.BackColor = System.Drawing.Color.Teal;
            this.excelFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.excelFile.Location = new System.Drawing.Point(1289, 1060);
            this.excelFile.Name = "excelFile";
            this.excelFile.Size = new System.Drawing.Size(356, 120);
            this.excelFile.TabIndex = 19;
            this.excelFile.Text = "Excel";
            this.excelFile.UseVisualStyleBackColor = false;
            this.excelFile.Click += new System.EventHandler(this.excelFile_Click);
            // 
            // readExcelButton
            // 
            this.readExcelButton.BackColor = System.Drawing.Color.Teal;
            this.readExcelButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.readExcelButton.Location = new System.Drawing.Point(799, 923);
            this.readExcelButton.Name = "readExcelButton";
            this.readExcelButton.Size = new System.Drawing.Size(419, 68);
            this.readExcelButton.TabIndex = 20;
            this.readExcelButton.Text = "Read + Refresh";
            this.readExcelButton.UseVisualStyleBackColor = false;
            this.readExcelButton.Click += new System.EventHandler(this.readExcelButton_Click);
            // 
            // listBox2
            // 
            this.listBox2.BackColor = System.Drawing.Color.LightCyan;
            this.listBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 29;
            this.listBox2.Location = new System.Drawing.Point(15, 97);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(311, 410);
            this.listBox2.TabIndex = 21;
            this.listBox2.SelectedIndexChanged += new System.EventHandler(this.listBox2_SelectedIndexChanged);
            // 
            // addAccountEventButton
            // 
            this.addAccountEventButton.BackColor = System.Drawing.Color.Teal;
            this.addAccountEventButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.addAccountEventButton.Location = new System.Drawing.Point(799, 1017);
            this.addAccountEventButton.Name = "addAccountEventButton";
            this.addAccountEventButton.Size = new System.Drawing.Size(419, 68);
            this.addAccountEventButton.TabIndex = 24;
            this.addAccountEventButton.Text = "Make Account Event";
            this.addAccountEventButton.UseVisualStyleBackColor = false;
            this.addAccountEventButton.Click += new System.EventHandler(this.addAccountEventButton_Click);
            // 
            // listBox3
            // 
            this.listBox3.BackColor = System.Drawing.Color.LightCyan;
            this.listBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox3.FormattingEnabled = true;
            this.listBox3.ItemHeight = 29;
            this.listBox3.Location = new System.Drawing.Point(407, 97);
            this.listBox3.Name = "listBox3";
            this.listBox3.Size = new System.Drawing.Size(311, 410);
            this.listBox3.TabIndex = 27;
            this.listBox3.SelectedIndexChanged += new System.EventHandler(this.listBox3_SelectedIndexChanged);
            // 
            // addTasksButton
            // 
            this.addTasksButton.BackColor = System.Drawing.Color.Teal;
            this.addTasksButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 21F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.addTasksButton.Location = new System.Drawing.Point(1289, 923);
            this.addTasksButton.Name = "addTasksButton";
            this.addTasksButton.Size = new System.Drawing.Size(356, 120);
            this.addTasksButton.TabIndex = 28;
            this.addTasksButton.Text = "Make All Account + Task Events";
            this.addTasksButton.UseVisualStyleBackColor = false;
            this.addTasksButton.Click += new System.EventHandler(this.addTasksButton_Click);
            // 
            // goToBedBox
            // 
            this.goToBedBox.BackColor = System.Drawing.Color.LightCyan;
            this.goToBedBox.FormattingEnabled = true;
            this.goToBedBox.Items.AddRange(new object[] {
            "00:00",
            "01:00",
            "02:00",
            "03:00",
            "04:00",
            "05:00",
            "06:00",
            "07:00",
            "08:00",
            "09:00",
            "10:00",
            "11:00",
            "12:00",
            "13:00",
            "14:00",
            "15:00",
            "16:00",
            "17:00",
            "18:00",
            "19:00",
            "20:00",
            "21:00",
            "22:00",
            "23:00",
            "24:00"});
            this.goToBedBox.Location = new System.Drawing.Point(1289, 452);
            this.goToBedBox.Name = "goToBedBox";
            this.goToBedBox.Size = new System.Drawing.Size(248, 28);
            this.goToBedBox.TabIndex = 29;
            // 
            // wakeUpBox
            // 
            this.wakeUpBox.BackColor = System.Drawing.Color.LightCyan;
            this.wakeUpBox.FormattingEnabled = true;
            this.wakeUpBox.Items.AddRange(new object[] {
            "00:00",
            "01:00",
            "02:00",
            "03:00",
            "04:00",
            "05:00",
            "06:00",
            "07:00",
            "08:00",
            "09:00",
            "10:00",
            "11:00",
            "12:00",
            "13:00",
            "14:00",
            "15:00",
            "16:00",
            "17:00",
            "18:00",
            "19:00",
            "20:00",
            "21:00",
            "22:00",
            "23:00",
            "24:00"});
            this.wakeUpBox.Location = new System.Drawing.Point(1289, 532);
            this.wakeUpBox.Name = "wakeUpBox";
            this.wakeUpBox.Size = new System.Drawing.Size(248, 28);
            this.wakeUpBox.TabIndex = 30;
            // 
            // doNotDisturbButton
            // 
            this.doNotDisturbButton.BackColor = System.Drawing.Color.Teal;
            this.doNotDisturbButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.doNotDisturbButton.Location = new System.Drawing.Point(1289, 624);
            this.doNotDisturbButton.Name = "doNotDisturbButton";
            this.doNotDisturbButton.Size = new System.Drawing.Size(356, 90);
            this.doNotDisturbButton.TabIndex = 31;
            this.doNotDisturbButton.Text = "Do Not Disturb";
            this.doNotDisturbButton.UseVisualStyleBackColor = false;
            this.doNotDisturbButton.Click += new System.EventHandler(this.doNotDisturbButton_Click);
            // 
            // taskBox
            // 
            this.taskBox.BackColor = System.Drawing.Color.LightCyan;
            this.taskBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.taskBox.Location = new System.Drawing.Point(407, 557);
            this.taskBox.Name = "taskBox";
            this.taskBox.Size = new System.Drawing.Size(311, 35);
            this.taskBox.TabIndex = 32;
            // 
            // monthCalendar2
            // 
            this.monthCalendar2.BackColor = System.Drawing.Color.LightCyan;
            this.monthCalendar2.Location = new System.Drawing.Point(407, 637);
            this.monthCalendar2.Name = "monthCalendar2";
            this.monthCalendar2.TabIndex = 33;
            this.monthCalendar2.TodayDate = new System.DateTime(((long)(0)));
            this.monthCalendar2.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar2_DateSelected);
            // 
            // addTaskButton
            // 
            this.addTaskButton.BackColor = System.Drawing.Color.Teal;
            this.addTaskButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.addTaskButton.Location = new System.Drawing.Point(407, 1218);
            this.addTaskButton.Name = "addTaskButton";
            this.addTaskButton.Size = new System.Drawing.Size(311, 108);
            this.addTaskButton.TabIndex = 34;
            this.addTaskButton.Text = "Add Task";
            this.addTaskButton.UseVisualStyleBackColor = false;
            this.addTaskButton.Click += new System.EventHandler(this.addTaskButton_Click);
            // 
            // accountListLabel
            // 
            this.accountListLabel.AutoSize = true;
            this.accountListLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 26F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.accountListLabel.Location = new System.Drawing.Point(19, 20);
            this.accountListLabel.Name = "accountListLabel";
            this.accountListLabel.Size = new System.Drawing.Size(307, 59);
            this.accountListLabel.TabIndex = 35;
            this.accountListLabel.Text = "Account List";
            // 
            // taskListLabel
            // 
            this.taskListLabel.AutoSize = true;
            this.taskListLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 26F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.taskListLabel.Location = new System.Drawing.Point(437, 20);
            this.taskListLabel.Name = "taskListLabel";
            this.taskListLabel.Size = new System.Drawing.Size(232, 59);
            this.taskListLabel.TabIndex = 36;
            this.taskListLabel.Text = "Task List";
            // 
            // taskNameLabel
            // 
            this.taskNameLabel.AutoSize = true;
            this.taskNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.taskNameLabel.Location = new System.Drawing.Point(401, 517);
            this.taskNameLabel.Name = "taskNameLabel";
            this.taskNameLabel.Size = new System.Drawing.Size(164, 36);
            this.taskNameLabel.TabIndex = 37;
            this.taskNameLabel.Text = "Task Name";
            // 
            // taskDateLabel
            // 
            this.taskDateLabel.AutoSize = true;
            this.taskDateLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.taskDateLabel.Location = new System.Drawing.Point(401, 597);
            this.taskDateLabel.Name = "taskDateLabel";
            this.taskDateLabel.Size = new System.Drawing.Size(76, 36);
            this.taskDateLabel.TabIndex = 38;
            this.taskDateLabel.Text = "Date";
            // 
            // eventListLabel
            // 
            this.eventListLabel.AutoSize = true;
            this.eventListLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 26F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.eventListLabel.Location = new System.Drawing.Point(867, 20);
            this.eventListLabel.Name = "eventListLabel";
            this.eventListLabel.Size = new System.Drawing.Size(252, 59);
            this.eventListLabel.TabIndex = 39;
            this.eventListLabel.Text = "Event List";
            // 
            // descriptionLabel
            // 
            this.descriptionLabel.AutoSize = true;
            this.descriptionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 26F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.descriptionLabel.Location = new System.Drawing.Point(1327, 20);
            this.descriptionLabel.Name = "descriptionLabel";
            this.descriptionLabel.Size = new System.Drawing.Size(284, 59);
            this.descriptionLabel.TabIndex = 40;
            this.descriptionLabel.Text = "Description";
            // 
            // goToSleepLabel
            // 
            this.goToSleepLabel.AutoSize = true;
            this.goToSleepLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.goToSleepLabel.Location = new System.Drawing.Point(1283, 412);
            this.goToSleepLabel.Name = "goToSleepLabel";
            this.goToSleepLabel.Size = new System.Drawing.Size(254, 36);
            this.goToSleepLabel.TabIndex = 41;
            this.goToSleepLabel.Text = "Go To Sleep Time";
            // 
            // wakeUpTimeLabel
            // 
            this.wakeUpTimeLabel.AutoSize = true;
            this.wakeUpTimeLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.wakeUpTimeLabel.Location = new System.Drawing.Point(1289, 492);
            this.wakeUpTimeLabel.Name = "wakeUpTimeLabel";
            this.wakeUpTimeLabel.Size = new System.Drawing.Size(210, 36);
            this.wakeUpTimeLabel.TabIndex = 42;
            this.wakeUpTimeLabel.Text = "Wake Up Time";
            // 
            // addTaskEventButton
            // 
            this.addTaskEventButton.BackColor = System.Drawing.Color.Teal;
            this.addTaskEventButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.addTaskEventButton.Location = new System.Drawing.Point(799, 1112);
            this.addTaskEventButton.Name = "addTaskEventButton";
            this.addTaskEventButton.Size = new System.Drawing.Size(419, 68);
            this.addTaskEventButton.TabIndex = 43;
            this.addTaskEventButton.Text = "Make Task Event";
            this.addTaskEventButton.UseVisualStyleBackColor = false;
            this.addTaskEventButton.Click += new System.EventHandler(this.addTaskEventButton_Click);
            // 
            // richTextBox2
            // 
            this.richTextBox2.Location = new System.Drawing.Point(1687, 97);
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.Size = new System.Drawing.Size(425, 536);
            this.richTextBox2.TabIndex = 44;
            this.richTextBox2.Text = "";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1289, 852);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(164, 65);
            this.button1.TabIndex = 45;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // maxStartBox
            // 
            this.maxStartBox.Location = new System.Drawing.Point(1289, 733);
            this.maxStartBox.Name = "maxStartBox";
            this.maxStartBox.Size = new System.Drawing.Size(210, 26);
            this.maxStartBox.TabIndex = 46;
            // 
            // maxCountBox
            // 
            this.maxCountBox.Location = new System.Drawing.Point(1289, 777);
            this.maxCountBox.Name = "maxCountBox";
            this.maxCountBox.Size = new System.Drawing.Size(210, 26);
            this.maxCountBox.TabIndex = 47;
            // 
            // richTextBox3
            // 
            this.richTextBox3.Location = new System.Drawing.Point(1687, 663);
            this.richTextBox3.Name = "richTextBox3";
            this.richTextBox3.Size = new System.Drawing.Size(425, 380);
            this.richTextBox3.TabIndex = 48;
            this.richTextBox3.Text = "";
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(1687, 1072);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(425, 93);
            this.button2.TabIndex = 49;
            this.button2.Text = "Find Date";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // monthCalendar3
            // 
            this.monthCalendar3.Location = new System.Drawing.Point(407, 907);
            this.monthCalendar3.Name = "monthCalendar3";
            this.monthCalendar3.TabIndex = 50;
            this.monthCalendar3.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar3_DateSelected);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(1687, 1183);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(425, 83);
            this.button3.TabIndex = 51;
            this.button3.Text = "Cancel Dates";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SteelBlue;
            this.ClientSize = new System.Drawing.Size(2141, 1465);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.monthCalendar3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.richTextBox3);
            this.Controls.Add(this.maxCountBox);
            this.Controls.Add(this.maxStartBox);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.richTextBox2);
            this.Controls.Add(this.addTaskEventButton);
            this.Controls.Add(this.wakeUpTimeLabel);
            this.Controls.Add(this.goToSleepLabel);
            this.Controls.Add(this.descriptionLabel);
            this.Controls.Add(this.eventListLabel);
            this.Controls.Add(this.taskDateLabel);
            this.Controls.Add(this.taskNameLabel);
            this.Controls.Add(this.taskListLabel);
            this.Controls.Add(this.accountListLabel);
            this.Controls.Add(this.addTaskButton);
            this.Controls.Add(this.monthCalendar2);
            this.Controls.Add(this.taskBox);
            this.Controls.Add(this.doNotDisturbButton);
            this.Controls.Add(this.wakeUpBox);
            this.Controls.Add(this.goToBedBox);
            this.Controls.Add(this.addTasksButton);
            this.Controls.Add(this.listBox3);
            this.Controls.Add(this.addAccountEventButton);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.readExcelButton);
            this.Controls.Add(this.excelFile);
            this.Controls.Add(this.accountMinPaymentLabel);
            this.Controls.Add(this.minPaymentBox);
            this.Controls.Add(this.accountBalanceLabel);
            this.Controls.Add(this.accountDueLabel);
            this.Controls.Add(this.accountNameLabel);
            this.Controls.Add(this.balanceBox);
            this.Controls.Add(this.displayDateBox);
            this.Controls.Add(this.accountBox);
            this.Controls.Add(this.addAccountButton);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.textBox1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Location = new System.Drawing.Point(42, 615);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Scheduler";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button addAccountButton;
        private System.Windows.Forms.TextBox accountBox;
        private System.Windows.Forms.TextBox displayDateBox;
        private System.Windows.Forms.TextBox balanceBox;
        private System.Windows.Forms.Label accountNameLabel;
        private System.Windows.Forms.Label accountDueLabel;
        private System.Windows.Forms.Label accountBalanceLabel;
        private System.Windows.Forms.TextBox minPaymentBox;
        private System.Windows.Forms.Label accountMinPaymentLabel;
        private System.Windows.Forms.Button excelFile;
        private System.Windows.Forms.Button readExcelButton;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Button addAccountEventButton;
        private System.Windows.Forms.ListBox listBox3;
        private System.Windows.Forms.Button addTasksButton;
        private System.Windows.Forms.ComboBox goToBedBox;
        private System.Windows.Forms.ComboBox wakeUpBox;
        private System.Windows.Forms.Button doNotDisturbButton;
        private System.Windows.Forms.TextBox taskBox;
        private System.Windows.Forms.MonthCalendar monthCalendar2;
        private System.Windows.Forms.Button addTaskButton;
        private System.Windows.Forms.Label accountListLabel;
        private System.Windows.Forms.Label taskListLabel;
        private System.Windows.Forms.Label taskNameLabel;
        private System.Windows.Forms.Label taskDateLabel;
        private System.Windows.Forms.Label eventListLabel;
        private System.Windows.Forms.Label descriptionLabel;
        private System.Windows.Forms.Label goToSleepLabel;
        private System.Windows.Forms.Label wakeUpTimeLabel;
        private System.Windows.Forms.Button addTaskEventButton;
        private System.Windows.Forms.RichTextBox richTextBox2;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox maxStartBox;
        private System.Windows.Forms.TextBox maxCountBox;
        private System.Windows.Forms.RichTextBox richTextBox3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.MonthCalendar monthCalendar3;
        private System.Windows.Forms.Button button3;
    }
}

