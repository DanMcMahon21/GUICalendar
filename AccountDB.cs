using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Runtime.InteropServices;

namespace GUICalendar
{
    public class AccountDB
    {
        // Read in all accounts from Excel database
        public static List<Account> ReadAccounts(List<Account> readAccounts)
        {
            string path = Path.GetDirectoryName(Directory.GetCurrentDirectory().ToString());
            //string excelFile = "C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs";
            string excelFile = "test505.xlxs";
            object missing = Type.Missing;

            Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;

            try
            {
                //Excel.Workbook wB = excelApp.Workbooks.Open(excelFile);
                Excel.Workbook wB = excelApp.Workbooks.Open($"{path}" + "\\" + $"{excelFile}");
                Excel.Sheets allSheets = wB.Worksheets;

                Excel.Worksheet wS1 = (Worksheet)allSheets["Accounts"];
                Excel.Worksheet wS2 = (Worksheet)allSheets["Tasks"];
                Excel.Worksheet wS3 = (Worksheet)allSheets["Events"];

                wS1.Activate();
                bool procede = true;
                int row = 2;
                int col = 1;
                int accountNum = 0;

                do
                {

                    if (wS1.Cells[row, col] == null || wS1.Cells[row, col].Value2 == null || wS1.Cells[row, col].Value2.ToString() == "")
                    {
                        procede = false;
                    }
                    else
                    {
                        Account a = new Account();
                        readAccounts.Add(a);
                        readAccounts[accountNum].Name = Convert.ToString(wS1.Cells[row, 1].Value2);
                        readAccounts[accountNum].Due = Convert.ToInt32(wS1.Cells[row, 2].Value2);
                        readAccounts[accountNum].Balance = Convert.ToDouble(wS1.Cells[row, 3].Value2);
                        readAccounts[accountNum].MinPayment = Convert.ToDouble(wS1.Cells[row, 4].Value2);
                        readAccounts[accountNum].DueDate = Convert.ToDateTime(wS1.Cells[row, 5].Value);
                        accountNum++;
                        row++;
                    }

                } while (procede);

                wB.Save();

                wB.Close();

                excelApp.Quit();


                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(allSheets);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);

            }
            catch
            {
                Excel._Workbook wB = null;

                wB = (Excel._Workbook)(excelApp.Workbooks.Add(Missing.Value));

                Excel.Worksheet wS1 = wB.ActiveSheet as Excel.Worksheet;
                wS1.Name = "Accounts";

                Excel.Worksheet wS2 = wB.Sheets.Add(missing, wS1, 1, missing) as Excel.Worksheet;
                wS2.Name = "Tasks";

                Excel.Worksheet wS3 = wB.Sheets.Add(missing, wS2, 1, missing) as Excel.Worksheet;
                wS3.Name = "Events";

                Excel.Range wR = null;
                Excel.Font wF = null;

                wR = wS1.Range[wS1.Cells[1, 1], wS1.Cells[1, 5]];
                wF = wR.Font;
                wF.Size = 15;

                wS1.Columns[1].ColumnWidth = 19;
                wS1.Columns[2].ColumnWidth = 12;
                wS1.Columns[3].ColumnWidth = 11;
                wS1.Columns[4].ColumnWidth = 23;
                wS1.Columns[5].ColumnWidth = 20;

                wS1.Cells[1, 1] = "Account Name";
                wS1.Cells[1, 2] = "Due Date";
                wS1.Cells[1, 3] = "Balance";
                wS1.Cells[1, 4] = "Minimum Payment";
                wS1.Cells[1, 5] = "DateTime Due";

                wR = wS2.Range[wS2.Cells[1, 1], wS2.Cells[1, 3]];
                wF = wR.Font;
                wF.Size = 15;

                wS2.Columns[1].ColumnWidth = 20;
                wS2.Columns[2].ColumnWidth = 12;
                wS2.Columns[3].ColumnWidth = 20;

                wS2.Cells[1, 1] = "Task Name";
                wS2.Cells[1, 2] = "Date";
                wS2.Cells[1, 3] = "DateTime Date";

                wR = wS3.Range[wS3.Cells[1, 1], wS3.Cells[1, 4]];
                wF = wR.Font;
                wF.Size = 15;

                wS3.Columns[1].ColumnWidth = 20;
                wS3.Columns[2].ColumnWidth = 12;
                wS3.Columns[3].ColumnWidth = 12;
                wS3.Columns[4].ColumnWidth = 20;

                wS3.Cells[1, 1] = "Event Name";
                wS3.Cells[1, 2] = "Start Time";
                wS3.Cells[1, 3] = "End Time";
                wS3.Cells[1, 4] = "Date Scheduled";

                //wB.SaveAs("C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs");
                wB.SaveAs($"{path}" + "\\" + $"{excelFile}");

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(wR);
                Marshal.ReleaseComObject(wF);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            return readAccounts;
        }

        // Read in all tasks from Excel database
        public static List<TaskN> ReadTasks(List<TaskN> readTasks)
        {
            string path = Path.GetDirectoryName(Directory.GetCurrentDirectory().ToString());
            //string excelFile = "C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs";
            string excelFile = "test505.xlxs";
            object missing = Type.Missing;

            Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;

            try
            {

                //Excel.Workbook wB = excelApp.Workbooks.Open(excelFile);
                Excel.Workbook wB = excelApp.Workbooks.Open($"{path}" + "\\" + $"{excelFile}");
                Excel.Sheets allSheets = wB.Worksheets;

                Excel.Worksheet wS1 = (Worksheet)allSheets["Accounts"];
                Excel.Worksheet wS2 = (Worksheet)allSheets["Tasks"];
                Excel.Worksheet wS3 = (Worksheet)allSheets["Events"];

                wS2.Activate();
                bool procede = true;
                int row = 2;
                int col = 1;
                int accountNum = 0;

                do
                {

                    if (wS2.Cells[row, col] == null || wS2.Cells[row, col].Value2 == null || wS2.Cells[row, col].Value2.ToString() == "")
                    {
                        procede = false;
                    }
                    else
                    {
                        TaskN b = new TaskN();
                        readTasks.Add(b);
                        readTasks[accountNum].Name = Convert.ToString(wS2.Cells[row, 1].Value2);
                        readTasks[accountNum].Date = Convert.ToInt32(wS2.Cells[row, 2].Value2);
                        readTasks[accountNum].TaskDate = Convert.ToDateTime(wS2.Cells[row, 3].Value);
                        accountNum++;
                        row++;
                    }

                } while (procede);

                wB.Save();

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(allSheets);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            catch
            {
                Excel._Workbook wB = null;

                wB = (Excel._Workbook)(excelApp.Workbooks.Add(Missing.Value));

                Excel.Worksheet wS1 = wB.ActiveSheet as Excel.Worksheet;
                wS1.Name = "Accounts";

                Excel.Worksheet wS2 = wB.Sheets.Add(missing, wS1, 1, missing) as Excel.Worksheet;
                wS2.Name = "Tasks";

                Excel.Worksheet wS3 = wB.Sheets.Add(missing, wS2, 1, missing) as Excel.Worksheet;
                wS3.Name = "Events";

                Excel.Range wR = null;
                Excel.Font wF = null;

                wR = wS1.Range[wS1.Cells[1, 1], wS1.Cells[1, 5]];
                wF = wR.Font;
                wF.Size = 15;

                wS1.Columns[1].ColumnWidth = 19;
                wS1.Columns[2].ColumnWidth = 12;
                wS1.Columns[3].ColumnWidth = 11;
                wS1.Columns[4].ColumnWidth = 23;
                wS1.Columns[5].ColumnWidth = 20;

                wS1.Cells[1, 1] = "Account Name";
                wS1.Cells[1, 2] = "Due Date";
                wS1.Cells[1, 3] = "Balance";
                wS1.Cells[1, 4] = "Minimum Payment";
                wS1.Cells[1, 5] = "DateTime Due";

                wR = wS2.Range[wS2.Cells[1, 1], wS2.Cells[1, 3]];
                wF = wR.Font;
                wF.Size = 15;

                wS2.Columns[1].ColumnWidth = 20;
                wS2.Columns[2].ColumnWidth = 12;
                wS2.Columns[3].ColumnWidth = 20;

                wS2.Cells[1, 1] = "Task Name";
                wS2.Cells[1, 2] = "Date";
                wS2.Cells[1, 3] = "DateTime Date";

                wR = wS3.Range[wS3.Cells[1, 1], wS3.Cells[1, 4]];
                wF = wR.Font;
                wF.Size = 15;

                wS3.Columns[1].ColumnWidth = 20;
                wS3.Columns[2].ColumnWidth = 12;
                wS3.Columns[3].ColumnWidth = 12;
                wS3.Columns[4].ColumnWidth = 20;

                wS3.Cells[1, 1] = "Event Name";
                wS3.Cells[1, 2] = "Start Time";
                wS3.Cells[1, 3] = "End Time";
                wS3.Cells[1, 4] = "Date Scheduled";

                //wB.SaveAs("C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs");
                wB.SaveAs($"{path}" + "\\" + $"{excelFile}");

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(wR);
                Marshal.ReleaseComObject(wF);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            return readTasks;
        }

        // Add new account to Excel database
        public static Account Add(Account account, int c)
        {
            string path = Path.GetDirectoryName(Directory.GetCurrentDirectory().ToString());
            //string excelFile = "C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs";
            string excelFile = "test505.xlxs";
            object missing = Type.Missing;

            Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;

            try
            {
                //Excel.Workbook wB = excelApp.Workbooks.Open(excelFile);
                Excel.Workbook wB = excelApp.Workbooks.Open($"{path}" + "\\" + $"{excelFile}");
                Excel.Sheets allSheets = wB.Worksheets;

                Excel.Worksheet wS1 = (Worksheet)allSheets["Accounts"];
                Excel.Worksheet wS2 = (Worksheet)allSheets["Tasks"];
                Excel.Worksheet wS3 = (Worksheet)allSheets["Events"];

                wS1.Activate();


                wS1.Cells[c + 1, 1] = $"{account.Name}";
                wS1.Cells[c + 1, 2] = $"{account.Due}";
                wS1.Cells[c + 1, 3] = $"{account.Balance}";
                wS1.Cells[c + 1, 4] = $"{account.MinPayment}";
                wS1.Cells[c + 1, 5] = $"{account.DueDate.ToString("d")}";

                wB.Save();

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(allSheets);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            catch
            {
                Excel._Workbook wB = null;

                wB = (Excel._Workbook)(excelApp.Workbooks.Add(Missing.Value));

                Excel.Worksheet wS1 = wB.ActiveSheet as Excel.Worksheet;
                wS1.Name = "Accounts";

                Excel.Worksheet wS2 = wB.Sheets.Add(missing, wS1, 1, missing) as Excel.Worksheet;
                wS2.Name = "Tasks";

                Excel.Worksheet wS3 = wB.Sheets.Add(missing, wS2, 1, missing) as Excel.Worksheet;
                wS3.Name = "Events";

                Excel.Range wR = null;
                Excel.Font wF = null;

                wR = wS1.Range[wS1.Cells[1, 1], wS1.Cells[1, 5]];
                wF = wR.Font;
                wF.Size = 15;

                wS1.Columns[1].ColumnWidth = 19;
                wS1.Columns[2].ColumnWidth = 12;
                wS1.Columns[3].ColumnWidth = 11;
                wS1.Columns[4].ColumnWidth = 23;
                wS1.Columns[5].ColumnWidth = 20;

                wS1.Cells[1, 1] = "Account Name";
                wS1.Cells[1, 2] = "Due Date";
                wS1.Cells[1, 3] = "Balance";
                wS1.Cells[1, 4] = "Minimum Payment";
                wS1.Cells[1, 5] = "DateTime Due";

                wR = wS2.Range[wS2.Cells[1, 1], wS2.Cells[1, 3]];
                wF = wR.Font;
                wF.Size = 15;

                wS2.Columns[1].ColumnWidth = 20;
                wS2.Columns[2].ColumnWidth = 12;
                wS2.Columns[3].ColumnWidth = 20;

                wS2.Cells[1, 1] = "Task Name";
                wS2.Cells[1, 2] = "Date";
                wS2.Cells[1, 3] = "DateTime Date";

                wR = wS3.Range[wS3.Cells[1, 1], wS3.Cells[1, 4]];
                wF = wR.Font;
                wF.Size = 15;

                wS3.Columns[1].ColumnWidth = 20;
                wS3.Columns[2].ColumnWidth = 12;
                wS3.Columns[3].ColumnWidth = 12;
                wS3.Columns[4].ColumnWidth = 20;

                wS3.Cells[1, 1] = "Event Name";
                wS3.Cells[1, 2] = "Start Time";
                wS3.Cells[1, 3] = "End Time";
                wS3.Cells[1, 4] = "Date Scheduled";

                wS1.Cells[c + 1, 1] = $"{account.Name}";
                wS1.Cells[c + 1, 2] = $"{account.Due}";
                wS1.Cells[c + 1, 3] = $"{account.Balance}";
                wS1.Cells[c + 1, 4] = $"{account.MinPayment}";
                wS1.Cells[c + 1, 5] = $"{account.DueDate.ToString("d")}";

                //wB.SaveAs("C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs");
                wB.SaveAs($"{path}" + "\\" + $"{excelFile}");

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(wR);
                Marshal.ReleaseComObject(wF);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            return account;
        }

        // Add new task to Excel database
        public static TaskN AddTask(TaskN task, int c)
        {
            string path = Path.GetDirectoryName(Directory.GetCurrentDirectory().ToString());
            //string excelFile = "C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs";
            string excelFile = "test505.xlxs";
            object missing = Type.Missing;

            Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;

            try
            {
                //Excel.Workbook wB = excelApp.Workbooks.Open(excelFile);
                Excel.Workbook wB = excelApp.Workbooks.Open($"{path}" + "\\" + $"{excelFile}");
                Excel.Sheets allSheets = wB.Worksheets;

                Excel.Worksheet wS1 = (Worksheet)allSheets["Accounts"];
                Excel.Worksheet wS2 = (Worksheet)allSheets["Tasks"];
                Excel.Worksheet wS3 = (Worksheet)allSheets["Events"];

                wS2.Activate();


                wS2.Cells[c + 1, 1] = $"{task.Name}";
                wS2.Cells[c + 1, 2] = $"{task.Date}";
                wS2.Cells[c + 1, 3] = $"{task.TaskDate.ToString("d")}";

                wB.Save();

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(allSheets);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            catch
            {
                Excel._Workbook wB = null;

                wB = (Excel._Workbook)(excelApp.Workbooks.Add(Missing.Value));

                Excel.Worksheet wS1 = wB.ActiveSheet as Excel.Worksheet;
                wS1.Name = "Accounts";

                Excel.Worksheet wS2 = wB.Sheets.Add(missing, wS1, 1, missing) as Excel.Worksheet;
                wS2.Name = "Tasks";

                Excel.Worksheet wS3 = wB.Sheets.Add(missing, wS2, 1, missing) as Excel.Worksheet;
                wS3.Name = "Events";

                Excel.Range wR = null;
                Excel.Font wF = null;

                wR = wS1.Range[wS1.Cells[1, 1], wS1.Cells[1, 5]];
                wF = wR.Font;
                wF.Size = 15;

                wS1.Columns[1].ColumnWidth = 19;
                wS1.Columns[2].ColumnWidth = 12;
                wS1.Columns[3].ColumnWidth = 11;
                wS1.Columns[4].ColumnWidth = 23;
                wS1.Columns[5].ColumnWidth = 20;

                wS1.Cells[1, 1] = "Account Name";
                wS1.Cells[1, 2] = "Due Date";
                wS1.Cells[1, 3] = "Balance";
                wS1.Cells[1, 4] = "Minimum Payment";
                wS1.Cells[1, 5] = "DateTime Due";

                wR = wS2.Range[wS2.Cells[1, 1], wS2.Cells[1, 3]];
                wF = wR.Font;
                wF.Size = 15;

                wS2.Columns[1].ColumnWidth = 20;
                wS2.Columns[2].ColumnWidth = 12;
                wS2.Columns[3].ColumnWidth = 20;

                wS2.Cells[1, 1] = "Task Name";
                wS2.Cells[1, 2] = "Date";
                wS2.Cells[1, 3] = "DateTime Date";

                wR = wS3.Range[wS3.Cells[1, 1], wS3.Cells[1, 4]];
                wF = wR.Font;
                wF.Size = 15;

                wS3.Columns[1].ColumnWidth = 20;
                wS3.Columns[2].ColumnWidth = 12;
                wS3.Columns[3].ColumnWidth = 12;
                wS3.Columns[4].ColumnWidth = 20;

                wS3.Cells[1, 1] = "Event Name";
                wS3.Cells[1, 2] = "Start Time";
                wS3.Cells[1, 3] = "End Time";
                wS3.Cells[1, 4] = "Date Scheduled";

                wS2.Cells[c + 1, 1] = $"{task.Name}";
                wS2.Cells[c + 1, 2] = $"{task.Date}";
                wS2.Cells[c + 1, 3] = $"{task.TaskDate.ToString("d")}";

                //wB.SaveAs("C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs");
                wB.SaveAs($"{path}" + "\\" + $"{excelFile}");

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(wR);
                Marshal.ReleaseComObject(wF);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            return task;
        }

        // Add events to Excel database
        public static void AddEvents(Events el)
        {
            string path = Path.GetDirectoryName(Directory.GetCurrentDirectory().ToString());
            //string excelFile = "C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs";
            string excelFile = "test505.xlxs";
            object missing = Type.Missing;

            Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;

            try
            {
                //Excel.Workbook wB = excelApp.Workbooks.Open(excelFile);
                Excel.Workbook wB = excelApp.Workbooks.Open($"{path}" + "\\" + $"{excelFile}");
                Excel.Sheets allSheets = wB.Worksheets;

                Excel.Worksheet wS1 = (Worksheet)allSheets["Accounts"];
                Excel.Worksheet wS2 = (Worksheet)allSheets["Tasks"];
                Excel.Worksheet wS3 = (Worksheet)allSheets["Events"];

                wS3.Activate();

                Excel.Range wR = null;
                Excel.Font wF = null;

                wS3.Cells.ClearContents();

                wR = wS3.Range[wS3.Cells[1, 1], wS3.Cells[1, 4]];
                wF = wR.Font;
                wF.Size = 15;

                wS3.Columns[1].ColumnWidth = 20;
                wS3.Columns[2].ColumnWidth = 12;
                wS3.Columns[3].ColumnWidth = 12;
                wS3.Columns[4].ColumnWidth = 20;

                wS3.Cells[1, 1] = "Event Name";
                wS3.Cells[1, 2] = "Start Time";
                wS3.Cells[1, 3] = "End Time";
                wS3.Cells[1, 4] = "Date Scheduled";

                for (int q = 1; q <= el.Items.Count; q++)
                {
                    if (el.Items[q - 1].Start.DateTimeRaw == null)
                    {
                        wS3.Cells[q + 1, 1] = $"{el.Items[q - 1].Summary}";
                        wS3.Cells[q + 1, 2] = "All Day Event";
                        wS3.Cells[q + 1, 3] = "All Day Event";
                        wS3.Cells[q + 1, 4] = $"{Convert.ToDateTime(el.Items[q - 1].Start.Date).ToString("d")}";
                    }
                    else
                    {
                        wS3.Cells[q + 1, 1] = $"{el.Items[q - 1].Summary}";
                        DateTime dts = DateTime.Parse($"{el.Items[q - 1].Start.DateTimeRaw}");
                        DateTime dte = DateTime.Parse($"{el.Items[q - 1].End.DateTimeRaw}");

                        //wS2.Cells[q, 2] = $"{el.Items[q - 1].Start.DateTime}";
                        wS3.Cells[q + 1, 2] = $"{dts.ToString("HH:mm")}";

                        //wS2.Cells[q, 3] = $"{el.Items[q - 1].End.DateTime}";
                        wS3.Cells[q + 1, 3] = $"{dte.ToString("HH:mm")}";

                        wS3.Cells[q + 1, 4] = $"{Convert.ToDateTime(el.Items[q - 1].Start.DateTimeRaw).ToString("d")}";
                    }
                }

                wB.Save();

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(allSheets);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            catch
            {
                Excel._Workbook wB = null;

                wB = (Excel._Workbook)(excelApp.Workbooks.Add(Missing.Value));

                Excel.Worksheet wS1 = wB.ActiveSheet as Excel.Worksheet;
                wS1.Name = "Accounts";

                Excel.Worksheet wS2 = wB.Sheets.Add(missing, wS1, 1, missing) as Excel.Worksheet;
                wS2.Name = "Tasks";

                Excel.Worksheet wS3 = wB.Sheets.Add(missing, wS2, 1, missing) as Excel.Worksheet;
                wS3.Name = "Events";

                Excel.Range wR = null;
                Excel.Font wF = null;

                wR = wS1.Range[wS1.Cells[1, 1], wS1.Cells[1, 5]];
                wF = wR.Font;
                wF.Size = 15;

                wS1.Columns[1].ColumnWidth = 19;
                wS1.Columns[2].ColumnWidth = 12;
                wS1.Columns[3].ColumnWidth = 11;
                wS1.Columns[4].ColumnWidth = 23;
                wS1.Columns[5].ColumnWidth = 20;

                wS1.Cells[1, 1] = "Account Name";
                wS1.Cells[1, 2] = "Due Date";
                wS1.Cells[1, 3] = "Balance";
                wS1.Cells[1, 4] = "Minimum Payment";
                wS1.Cells[1, 5] = "DateTime Due";

                wR = wS2.Range[wS2.Cells[1, 1], wS2.Cells[1, 3]];
                wF = wR.Font;
                wF.Size = 15;

                wS2.Columns[1].ColumnWidth = 20;
                wS2.Columns[2].ColumnWidth = 12;
                wS2.Columns[3].ColumnWidth = 20;

                wS2.Cells[1, 1] = "Task Name";
                wS2.Cells[1, 2] = "Date";
                wS2.Cells[1, 3] = "DateTime Date";

                wR = wS3.Range[wS3.Cells[1, 1], wS3.Cells[1, 4]];
                wF = wR.Font;
                wF.Size = 15;

                wS3.Columns[1].ColumnWidth = 20;
                wS3.Columns[2].ColumnWidth = 12;
                wS3.Columns[3].ColumnWidth = 12;
                wS3.Columns[4].ColumnWidth = 20;

                wS3.Cells[1, 1] = "Event Name";
                wS3.Cells[1, 2] = "Start Time";
                wS3.Cells[1, 3] = "End Time";
                wS3.Cells[1, 4] = "Date Scheduled";

                wS3.Activate();

                for (int q = 1; q <= el.Items.Count; q++)
                {
                    if (el.Items[q - 1].Start.DateTimeRaw == null)
                    {
                        wS3.Cells[q + 1, 1] = $"{el.Items[q - 1].Summary}";
                        wS3.Cells[q + 1, 2] = "All Day Event";
                        wS3.Cells[q + 1, 3] = "All Day Event";
                        wS3.Cells[q + 1, 4] = $"{Convert.ToDateTime(el.Items[q - 1].Start.Date).ToString("d")}";
                    }
                    else
                    {
                        wS3.Cells[q + 1, 1] = $"{el.Items[q - 1].Summary}";
                        DateTime dts = DateTime.Parse($"{el.Items[q - 1].Start.DateTimeRaw}");
                        DateTime dte = DateTime.Parse($"{el.Items[q - 1].End.DateTimeRaw}");

                        //wS2.Cells[q, 2] = $"{el.Items[q - 1].Start.DateTime}";
                        wS3.Cells[q + 1, 2] = $"{dts.ToString("HH:mm")}";

                        //wS2.Cells[q, 3] = $"{el.Items[q - 1].End.DateTime}";
                        wS3.Cells[q + 1, 3] = $"{dte.ToString("HH:mm")}";

                        wS3.Cells[q + 1, 4] = $"{Convert.ToDateTime(el.Items[q - 1].Start.DateTimeRaw).ToString("d")}";
                    }
                }

                //wB.SaveAs("C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs");
                wB.SaveAs($"{path}" + "\\" + $"{excelFile}");

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(wR);
                Marshal.ReleaseComObject(wF);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
        }

        // Delete account from Excel database
        public static void DeleteAccount(int accountCount, int deleteIndex)
        {
            string path = Path.GetDirectoryName(Directory.GetCurrentDirectory().ToString());
            //string excelFile = "C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs";
            string excelFile = "test505.xlxs";
            object missing = Type.Missing;

            Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;

            try
            {
                //Excel.Workbook wB = excelApp.Workbooks.Open(excelFile);
                Excel.Workbook wB = excelApp.Workbooks.Open($"{path}" + "\\" + $"{excelFile}");
                Excel.Sheets allSheets = wB.Worksheets;

                Excel.Worksheet wS1 = (Worksheet)allSheets["Accounts"];
                Excel.Worksheet wS2 = (Worksheet)allSheets["Tasks"];
                Excel.Worksheet wS3 = (Worksheet)allSheets["Events"];

                wS2.Activate();

                for (int d = deleteIndex; d <= accountCount; d++)
                {
                    if (d < accountCount - 1)
                    {
                        wS1.Cells[d + 2, 1] = $"{Convert.ToString(wS1.Cells[d + 3, 1].Value2)}";
                        wS1.Cells[d + 2, 2] = $"{Convert.ToInt32(wS1.Cells[d + 3, 2].Value2)}";
                        wS1.Cells[d + 2, 3] = $"{Convert.ToDouble(wS1.Cells[d + 3, 3].Value2)}";
                        wS1.Cells[d + 2, 4] = $"{Convert.ToDouble(wS1.Cells[d + 3, 4].Value2)}";
                        wS1.Cells[d + 2, 5] = $"{Convert.ToDateTime(wS1.Cells[d + 3, 5].Value)}";
                    }
                    else
                    {
                        wS1.Cells[d + 2, 1].ClearContents();
                        wS1.Cells[d + 2, 2].ClearContents();
                        wS1.Cells[d + 2, 3].ClearContents();
                        wS1.Cells[d + 2, 4].ClearContents();
                        wS1.Cells[d + 2, 5].ClearContents();

                        wS1.Cells[d + 2, 1].ClearFormats();
                        wS1.Cells[d + 2, 2].ClearFormats();
                        wS1.Cells[d + 2, 3].ClearFormats();
                        wS1.Cells[d + 2, 4].ClearFormats();
                        wS1.Cells[d + 2, 5].ClearFormats();
                    }
                }

                wB.Save();

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(allSheets);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            catch
            {
                excelApp.Quit();
                MessageBox.Show("Counld Not Delete Account");
                Marshal.ReleaseComObject(excelApp);
            }
        }

        // Delete task from Excel database
        public static void DeleteTask(int taskCount, int deleteIndex)
        {
            string path = Path.GetDirectoryName(Directory.GetCurrentDirectory().ToString());
            //string excelFile = "C:\\Users\\dpmja\\Desktop\\Dan\\Dan School Files\\Grad School\\Final Grad Project\\GUICalendar\\test505.xlxs";
            string excelFile = "test505.xlxs";
            object missing = Type.Missing;

            Excel.Application excelApp = new Excel.Application();
            //excelApp.Visible = true;

            try
            {
                //Excel.Workbook wB = excelApp.Workbooks.Open(excelFile);
                Excel.Workbook wB = excelApp.Workbooks.Open($"{path}" + "\\" + $"{excelFile}");
                Excel.Sheets allSheets = wB.Worksheets;

                Excel.Worksheet wS1 = (Worksheet)allSheets["Accounts"];
                Excel.Worksheet wS2 = (Worksheet)allSheets["Tasks"];
                Excel.Worksheet wS3 = (Worksheet)allSheets["Events"];

                wS2.Activate();

                for(int d = deleteIndex; d <= taskCount; d++)
                {
                    if(d < taskCount - 1)
                    {
                        wS2.Cells[d + 2, 1] = $"{Convert.ToString(wS2.Cells[d + 3, 1].Value2)}";
                        wS2.Cells[d + 2, 2] = $"{Convert.ToInt32(wS2.Cells[d + 3, 2].Value2)}";
                        wS2.Cells[d + 2, 3] = $"{Convert.ToDateTime(wS2.Cells[d + 3, 3].Value)}";
                    }
                    else
                    {
                        wS2.Cells[d + 2, 1].ClearContents();
                        wS2.Cells[d + 2, 2].ClearContents();
                        wS2.Cells[d + 2, 3].ClearContents();

                        wS2.Cells[d + 2, 1].ClearFormats();
                        wS2.Cells[d + 2, 2].ClearFormats();
                        wS2.Cells[d + 2, 3].ClearFormats();
                    }
                }

                wB.Save();

                wB.Close();

                excelApp.Quit();

                Marshal.ReleaseComObject(wB);
                Marshal.ReleaseComObject(allSheets);
                Marshal.ReleaseComObject(wS1);
                Marshal.ReleaseComObject(wS2);
                Marshal.ReleaseComObject(wS3);
                Marshal.ReleaseComObject(excelApp);
            }
            catch
            {
                excelApp.Quit();
                MessageBox.Show("Counld Not Delete Task");
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
