using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ExcelSpliter3
{
    public partial class MainForm : Form
    {
        private string ExcelFile = "";

        public MainForm()
        {
            InitializeComponent();
            //Instantiate an instance of license and set the license file through its path

            Aspose.Cells.License license = new Aspose.Cells.License();

            license.SetLicense("License.txt");
            /*            string LData = "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz4NCjxMaWNlbnNlPg0KICAgIDxEYXRhPg0KICAgICAgICA8TGljZW5zZWRUbz5pckRldmVsb3BlcnMuY29tPC9MaWNlbnNlZFRvPg0KICAgICAgICA8RW1haWxUbz5pbmZvQGlyRGV2ZWxvcGVycy5jb208L0VtYWlsVG8+DQogICAgICAgIDxMaWNlbnNlVHlwZT5EZXZlbG9wZXIgT0VNPC9MaWNlbnNlVHlwZT4NCiAgICAgICAgPExpY2Vuc2VOb3RlPkxpbWl0ZWQgdG8gMTAwMCBkZXZlbG9wZXIsIHVubGltaXRlZCBwaHlzaWNhbCBsb2NhdGlvbnM8L0xpY2Vuc2VOb3RlPg0KICAgICAgICA8T3JkZXJJRD43ODQzMzY0Nzc4NTwvT3JkZXJJRD4NCiAgICAgICAgPFVzZXJJRD4xMTk0NDkyNDM3OTwvVXNlcklEPg0KICAgICAgICA8T0VNPlRoaXMgaXMgYSByZWRpc3RyaWJ1dGFibGUgbGljZW5zZTwvT0VNPg0KICAgICAgICA8UHJvZHVjdHM+DQogICAgICAgICAgICA8UHJvZHVjdD5Bc3Bvc2UuVG90YWwgUHJvZHVjdCBGYW1pbHk8L1Byb2R1Y3Q+DQogICAgICAgIDwvUHJvZHVjdHM+DQogICAgICAgIDxFZGl0aW9uVHlwZT5FbnRlcnByaXNlPC9FZGl0aW9uVHlwZT4NCiAgICAgICAgPFNlcmlhbE51bWJlcj57RjJCOTcwNDUtMUIyOS00QjNGLUJENTMtNjAxRUZGQTE1QUE5fTwvU2VyaWFsTnVtYmVyPg0KICAgICAgICA8U3Vic2NyaXB0aW9uRXhwaXJ5PjIwOTkxMjMxPC9TdWJzY3JpcHRpb25FeHBpcnk+DQogICAgICAgIDxMaWNlbnNlVmVyc2lvbj4zLjA8L0xpY2Vuc2VWZXJzaW9uPg0KICAgIDwvRGF0YT4NCiAgICA8U2lnbmF0dXJlPlFYTndiM05sTGxSdmRHRnNMb1B5YjJSMVkzUWdSbUZ0YVd4NTwvU2lnbmF0dXJlPg0KPC9MaWNlbnNlPg==";

                        Stream stream = new MemoryStream(Convert.FromBase64String(LData));

                        stream.Seek(0, SeekOrigin.Begin);

                        new Aspose.Cells.License().SetLicense(stream);
                        */
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            PopulateCheckedListBoxFromSaveFormatEnum(ExportTo);
            UpdateExportControls();
        }

        public void PopulateCheckedListBoxFromSaveFormatEnum(CheckedListBox checkedListBox)
        {
            // Get the names and values of the SaveFormat enum's members
            string[] names = Enum.GetNames(typeof(SaveFormat));
            Array values = Enum.GetValues(typeof(SaveFormat));

            // Add each member to the CheckedListBox with its name as the display text
            for (int i = 0; i < values.Length; i++)
            {
                checkedListBox.Items.Add(names[i], false);
            }
        }

        private void logOld(string log)
        {
            logBox.Items.Add(log);
            logBox.Refresh();
            logBox.SelectedIndex = logBox.Items.Count - 1;
        }

        public void log(string logMessage)
        {
            if (this.logBox.InvokeRequired)
            {
                this.logBox.Invoke(new Action<string>(log), new object[] { logMessage });
            }
            else
            {
                this.logBox.Items.Add(logMessage);
            }
        }

        private void LoadDataIntoGrid(string fileName, string sheetName, DataGridView dataGridView)
        {
            // Load the Excel file
            Workbook workbook = new Workbook(fileName);

            // Get the selected worksheet
            Worksheet worksheet = workbook.Worksheets[sheetName];

            // Create a DataTable to hold the data
            DataTable dataTable = new DataTable();

            // Add columns to the DataTable based on the header row
            for (int column = 0; column <= worksheet.Cells.MaxDataColumn; column++)
            {
                dataTable.Columns.Add(worksheet.Cells[0, column].StringValue);
            }

            // Add rows to the DataTable
            int rowCount = Math.Min(10, worksheet.Cells.MaxDataRow + 1); // Get the minimum of 10 or the total number of rows
            for (int row = 1; row <= rowCount; row++)
            {
                DataRow dataRow = dataTable.Rows.Add();
                for (int column = 0; column <= worksheet.Cells.MaxDataColumn; column++)
                {
                    dataRow[column] = worksheet.Cells[row, column].Value;
                }
            }

            // Bind the DataTable to the DataGrid
            dataGridView.DataSource = dataTable;
        }

        private void FillComboBoxWithSheetNames(string fileName, ComboBox comboBox)
        {
            // Load the workbook
            Workbook workbook = new Workbook(fileName);

            // Clear the existing items in the ComboBox
            comboBox.Items.Clear();

            // Add the sheet names to the ComboBox
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                comboBox.Items.Add(worksheet.Name);
            }

            // Select the first sheet by default
            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
            }
        }

        private void FillComboBoxWithColumnHeaders(string fileName, string selectedSheet, ComboBox comboBox)
        {
            // Load the workbook
            Workbook workbook = new Workbook(fileName);

            // Get the selected worksheet
            Worksheet worksheet = workbook.Worksheets[selectedSheet];

            // Clear the existing items in the ComboBox
            comboBox.Items.Clear();

            // Add the column headers to the ComboBox
            foreach (Cell cell in worksheet.Cells.Rows[0])
            {
                string columnHeader = cell.StringValue;
                comboBox.Items.Add(columnHeader);
            }

            // Select the first column header by default
            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
            }
        }


        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception occurred while releasing object: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void SendEmail(string toAddress, string subject, string attachmentFilePath, string body)
        {
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                mailItem.Subject = subject;
                mailItem.Body = body;
                mailItem.To = toAddress;

                if (!string.IsNullOrEmpty(attachmentFilePath))
                {
                    mailItem.Attachments.Add(attachmentFilePath);
                }

                //mailItem.Display(); // Display the email in Outlook for the user to review or send manually

                // If you want to send the email automatically, uncomment the following lines and comment out the "mailItem.Display()" line
                mailItem.Send();

                ReleaseObject(mailItem);
                ReleaseObject(outlookApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while sending the email: " + ex.Message);
            }
        }

        public void SplitSheetByColumnValue(string fileName, string worksheetName, string columnName)
        {
            // Load the Excel file
            Workbook workbook = new Workbook(fileName);

            // Get the specified worksheet
            Worksheet worksheet = workbook.Worksheets[worksheetName];
            

            // Find the column index based on the column name
            Cells cells = worksheet.Cells;
            int columnIndex = -1;
            for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
            {
                if (cells[0, col].StringValue == columnName)
                {
                    columnIndex = col;
                    break;
                }
            }

            if (columnIndex == -1)
            {
                // Column name not found
                return;
            }

            // Get the unique values in the specified column
            //Range columnRange = cells.CreateRange(cells.MinDataRow + 1, columnIndex, cells.MaxDataRow, columnIndex);
            Range columnRange = cells.CreateRange(cells.MinDataRow, columnIndex, cells.MaxDataRow - cells.MinDataRow + 1, 1);

            object[,] columnValues = columnRange.Value as object[,];
            object[] uniqueValues = columnValues.Cast<object>().Distinct().ToArray();

            // Split the sheet into multiple files based on the unique column values
            for (int i = 1; i < uniqueValues.Length; i++)
            {
                string columnValue = uniqueValues[i].ToString();

                // Create a new workbook and worksheet
                Workbook newWorkbook = new Workbook();
                newWorkbook.BuiltInDocumentProperties.Author = "شهاب صادقی";
                newWorkbook.BuiltInDocumentProperties.Comments = "این فایل توسط نرم افزار اکسل اسپلیتر 3 ایجاد شده است";
                if (setWorkbookPassword.Checked)
                {
                    //newWorkbook.Protect(ProtectionType.Structure, workBookPassword.Text);
                    newWorkbook.Settings.WriteProtection.Password = workBookPassword.Text;
                    newWorkbook.Settings.WriteProtection.Author = "شهاب صادقی";
                }                    
                    
                    
                    
                Worksheet newWorksheet = newWorkbook.Worksheets[0];

                // Copy the header row
                for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
                {
                    newWorksheet.Cells[0, col].Copy(cells[0, col]);
                }

                // Copy the rows for the specific column value
                int newRow = 1;
                for (int row = cells.MinDataRow + 1; row <= cells.MaxDataRow; row++)
                {
                    if (cells[row, columnIndex].StringValue == columnValue)
                    {
                        for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
                        {
                            newWorksheet.Cells[newRow, col].Copy(cells[row, col]);
                        }
                        newRow++;
                    }
                }


                if (RemoveDuplicates.Checked)
                    newWorksheet.Cells.RemoveDuplicates();



                // Remove the remaining empty rows
                newWorksheet.Cells.DeleteRows(newRow, newWorksheet.Cells.MaxDataRow - newRow + 1, true);
                if (Autofit.Checked)
                {
                    newWorksheet.AutoFitColumns();
                }

                if (freezTopRow.Checked)
                {
                    try
                    {
                        newWorksheet.FreezePanes(1, 0, 1, 0);
                    }
                    catch (Exception e)
                    {
                        log("خطا در فریز کردن ردیف اول");
                    }
                   
                }
                
                // Save the new workbook to a separate file
                string newFileName = Path.Combine(Path.GetDirectoryName(fileName), $"{RemoveInvalidFileNameChars(columnValue)}_{Path.GetFileNameWithoutExtension(fileName)}");

                foreach (var format in ExportTo.CheckedItems)
                {
                    SaveFormat formattoSave = (SaveFormat)Enum.Parse(typeof(SaveFormat),format.ToString(),true);
                    newWorkbook.Save(newFileName+"."+format,formattoSave);
                    log(newFileName+"."+format);
                    if (sendMail.Checked)
                    {
                        SendEmail(columnValue + MailPart.Text, mailSubject.Text, newFileName + "." + format, mailBody.Text);
                        log("Mail to " + columnValue);
                    }
                }
                
                


            }
        }


        public void SplitSheetByColumnValueAsync(string fileName, string worksheetName, string columnName)
        {
            // Load the Excel file
            Workbook workbook = new Workbook(fileName);

            // Get the specified worksheet
            Worksheet worksheet = workbook.Worksheets[worksheetName];

            // Find the column index based on the column name
            Cells cells = worksheet.Cells;
            int columnIndex = -1;
            for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
            {
                if (cells[0, col].StringValue == columnName)
                {
                    columnIndex = col;
                    break;
                }
            }

            if (columnIndex == -1)
            {
                // Column name not found
                return;
            }

            // Get the unique values in the specified column
            Range columnRange = cells.CreateRange(cells.MinDataRow, columnIndex, cells.MaxDataRow - cells.MinDataRow + 1, 1);

            object[,] columnValues = columnRange.Value as object[,];
            object[] uniqueValues = columnValues.Cast<object>().Distinct().ToArray();

            // Split the sheet into multiple files based on the unique column values
            Parallel.For(1, uniqueValues.Length, i =>
            {
                string columnValue = uniqueValues[i].ToString();

                // Create a new workbook and worksheet
                Workbook newWorkbook = new Workbook();
                newWorkbook.BuiltInDocumentProperties.Author = "شهاب صادقی";
                newWorkbook.BuiltInDocumentProperties.Comments = "این فایل توسط نرم افزار اکسل اسپلیتر 3 ایجاد شده است";
                if (setWorkbookPassword.Checked)
                {
                    newWorkbook.Settings.WriteProtection.Password = workBookPassword.Text;
                    newWorkbook.Settings.WriteProtection.Author = "شهاب صادقی";
                }

                Worksheet newWorksheet = newWorkbook.Worksheets[0];

                // Copy the header row
                for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
                {
                    newWorksheet.Cells[0, col].Copy(cells[0, col]);
                }

                // Copy the rows for the specific column value
                int newRow = 1;
                for (int row = cells.MinDataRow + 1; row <= cells.MaxDataRow; row++)
                {
                    if (cells[row, columnIndex].StringValue == columnValue)
                    {
                        for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
                        {
                            newWorksheet.Cells[newRow, col].Copy(cells[row, col]);
                        }
                        newRow++;
                    }
                }

                if (RemoveDuplicates.Checked)
                    newWorksheet.Cells.RemoveDuplicates();

                // Remove the remaining empty rows
                newWorksheet.Cells.DeleteRows(newRow, newWorksheet.Cells.MaxDataRow - newRow + 1, true);
                if (Autofit.Checked)
                {
                    newWorksheet.AutoFitColumns();
                }

                if (freezTopRow.Checked)
                {
                    try
                    {
                        newWorksheet.FreezePanes(1, 0, 1, 0);
                    }
                    catch (Exception e)
                    {
                        log("خطا در فریز کردن ردیف اول");
                    }
                }

                // Save the new workbook to a separate file
                string newFileName = Path.Combine(Path.GetDirectoryName(fileName), $"{RemoveInvalidFileNameChars(columnValue)}_{Path.GetFileNameWithoutExtension(fileName)}");

                foreach (var format in ExportTo.CheckedItems)
                {
                    SaveFormat formattoSave = (SaveFormat)Enum.Parse(typeof(SaveFormat), format.ToString(), true);
                    newWorkbook.Save(newFileName + "." + format, formattoSave);
                    log(newFileName + "." + format);
                    if (sendMail.Checked)
                    {
                        SendEmail(columnValue + MailPart.Text, mailSubject.Text, newFileName + "." + format, mailBody.Text);
                        log("Mail to " + columnValue);
                    }
                }
            });
        }


        public static string RemoveInvalidFileNameChars(string input)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return new string(input.Where(c => !invalidChars.Contains(c)).ToArray());
        }

        public void SplitExcelFile(string filePath)
        {
            // Load the Excel file
            Workbook workbook = new Workbook(filePath);

            // Loop through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Create a new workbook and add the current worksheet to it
                Workbook newWorkbook = new Workbook();
                newWorkbook.BuiltInDocumentProperties.Author = "شهاب صادقی";
                newWorkbook.BuiltInDocumentProperties.Comments = "این فایل توسط نرم افزار اکسل اسپلیتر 3 ایجاد شده است";

                if (setWorkbookPassword.Checked)
                {
                    //newWorkbook.Protect(ProtectionType.Structure, workBookPassword.Text);
                    newWorkbook.Settings.WriteProtection.Password = workBookPassword.Text;
                    newWorkbook.Settings.WriteProtection.Author = "شهاب صادقی";
                }


                Worksheet newSheet = newWorkbook.Worksheets[0];
                newSheet.Copy(sheet);
                if (Autofit.Checked)
                {
                    newSheet.AutoFitColumns();
                }
                if (freezTopRow.Checked) newWorkbook.Worksheets[0].FreezePanes(1, 0, 1, 0);

                // Save the new workbook to a separate file
                string newFileName = Path.Combine(Path.GetDirectoryName(filePath), $"{sheet.Name}_{Path.GetFileNameWithoutExtension(filePath)}");

                foreach (var format in ExportTo.CheckedItems)
                {
                    SaveFormat formattoSave = (SaveFormat)Enum.Parse(typeof(SaveFormat), format.ToString(), true);
                    newWorkbook.Save(newFileName + "." + format, formattoSave);
                    log("Exported "+newFileName + "." + format);
                    if (sendMail.Checked)
                    {
                        SendEmail(sheet.Name + MailPart.Text, mailSubject.Text, newFileName + "." + format, mailBody.Text);
                        log("mail to " + sheet.Name);
                    }
                }

                

                
            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            // Create an instance of the OpenFileDialog class
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the filter and title properties
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog.Title = "Select an Excel File";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ExcelFile = openFileDialog.FileName;
                this.Text = "Excel Splitter 3 - " + openFileDialog.FileName;
                log("بارگیری " + ExcelFile);
                log("تهیه لیست کاربرگ ها" + ExcelFile);
                FillComboBoxWithSheetNames(ExcelFile, Worksheets);

                #region Enable Controls

                Worksheets.Enabled = true;
                Headers.Enabled = true;
                Sheet2File.Enabled = false;
                Sheets2Files.Enabled = true;
                sendMail.Enabled = true;
                ExportTo.Enabled = true;
                #endregion Enable Controls

            }
        }

        private void Worksheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            log("بارگیری عنوان ستون ها");

            FillComboBoxWithColumnHeaders(ExcelFile, Worksheets.SelectedItem.ToString(), Headers);

            log("بارگیری پیش نمایش");

            LoadDataIntoGrid(ExcelFile, Worksheets.SelectedItem.ToString(), dataGridView1);
        }

        private void Sheets2Files_Click(object sender, EventArgs e)
        {
            log("شروع تقسیم...");
            SplitExcelFile(ExcelFile);
            log("اتمام عملیات");
        }

        private void Sheet2File_Click(object sender, EventArgs e)
        {
            log("شروع تقسیم...");
            SplitSheetByColumnValue(ExcelFile, Worksheets.SelectedItem.ToString(), Headers.SelectedItem.ToString());
            log("اتمام عملیات");
        }

        private void V3_Click(object sender, EventArgs e)
        {
        }

        private void UpdateExportControls()
        {
            if (ExcelFile == null)
            {
                Sheets2Files.Enabled = false;
                Sheet2File.Enabled = false;
            }
            if (ExportTo.CheckedItems.Count > 1)
            {
                sendMail.Checked = false;
            }

            if (ExportTo.CheckedItems.Count == 0)
            {
                Sheets2Files.Enabled = false;
                Sheet2File.Enabled = false;
            }
            else if (ExportTo.CheckedItems.Count == 1)
            {
                Sheets2Files.Enabled = true;
                Sheet2File.Enabled = true;
            }
        }

        private void ExportTo_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            UpdateExportControls();
        }

        private void ExportTo_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateExportControls();
        }

        private void sendMail_CheckedChanged(object sender, EventArgs e)
        {
            UpdateExportControls();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            log("شروع تقسیم آسینک...");
            SplitSheetByColumnValueAsync(ExcelFile, Worksheets.SelectedItem.ToString(), Headers.SelectedItem.ToString());
            log("اتمام عملیات");
        }
    }
}