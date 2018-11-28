using Microsoft.Office.Interop.Excel;
using ModbusUber;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using TestExcelSolar.Model;

namespace TestExcelSolar
{
    class Program
    {


        static void Main(string[] args)
        {
            string connString = "";
            string ExcelFilePath = "C:\\Users\\Anbarasu\\Desktop\\InvertorData.xlsx";
            string temp = Path.GetFileName(ExcelFilePath).ToLower();
            int _solarSerialNo = 0;
            if (temp.Trim() == ".xls")//Connection String to Excel Workbook
            {
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ExcelFilePath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (temp.Trim() == ".xlsx")
            {
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }
            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            string query = "Select * from [Sheet1$]";
            OleDbConnection conn = new OleDbConnection(connString);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            OleDbCommand cmd = new OleDbCommand(query, conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            System.Data.DataTable dataTable = new System.Data.DataTable();
            da.Fill(dataTable);
            //grvExcelData.DataSource = ds.Tables[0];
            //grvExcelData.DataBind();
            da.Dispose();
            conn.Close();
            conn.Dispose();
            int row = 4;
            foreach (DataRow datarowItem in dataTable.Rows)
            {
                var _houseNo = datarowItem.Field<string>("HouseNo");
                var _ipAddress = datarowItem.Field<string>("InverterIP");
                var _port = Convert.ToInt32(datarowItem.Field<double>("Port"));
                var _deviceID = Convert.ToInt32(datarowItem.Field<double>("DeviceID"));
                var _startAddress = Convert.ToInt32(datarowItem.Field<double>("StartAddress"));
                var _qty = Convert.ToInt32(datarowItem.Field<double>("length"));
                var _regType = Convert.ToInt32(datarowItem.Field<double>("registerType"));
                ReadExcel();
                Testy();
                Test();
                adfsdf();
                int[] readHoldingRegisters = ModbusReading.ReadRegisterWithDeviceIDs(_ipAddress, _port, _startAddress, _regType, _qty, Convert.ToByte(_deviceID));
                _solarSerialNo = GetMSB(readHoldingRegisters);
                //AddData(_houseNo, _ipAddress, _port, _solarSerialNo, 0, row);
                AddData(_houseNo, _ipAddress, _port, 1231231321, 0, row);
                row++;
            }

        }

        public static void adfsdf()
        {
            //      Microsoft.Office.Interop.Excel.Application app = null;
            //Microsoft.Office.Interop.Excel.Workbook workbook = null;
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
            //Microsoft.Office.Interop.Excel.Range workSheet_range = null;
            //const int FIRTSCOLUMN = 0; //Here const you will use to select good column
            //const int FIRSTROW = 0;
            //const int FIRSTSHEET = 1;
            //   app = new Microsoft.Office.Interop.Excel.Application();
            //   app.Visible = true;
            //   workbook = app.Workbooks.Add(1);
            //   worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[FIRSTSHEET];
            //   //addData(FIRSTROW, FIRTSCOLUMN, "yourdata");
            //   worksheet.Cells[FIRSTROW, FIRTSCOLUMN] = "asdsdfsdf";


            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            string myPath = CreateDirectory();
            //string myPath = @"Data.xlsx";
            excelApp.Workbooks.Open("C:\\Users\\Anbarasu\\Desktop\\yuiyuiyuiyui.xls");

            // Get Worksheet
            Microsoft.Office.Interop.Excel.Worksheet worksheet = excelApp.Worksheets[1];
            int rowIndex = 2; int colIndex = 2;
            for (int i = 0; i < 10; i++)
            {
                excelApp.Cells[rowIndex, colIndex] = "\r123dghghfdgjhfdgj";
            }

            excelApp.Visible = false;


            excelApp.ThisWorkbook.SaveAs("C:\\Users\\Anbarasu\\Desktop\\yuiyuiyuiyui.xls" + "fsdfsdf" + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            excelApp.ThisWorkbook.Close();
            Marshal.ReleaseComObject(excelApp);

            //string connectionString = "";
            //string fileName = @"C:\Users\\Anbarasu\Desktop\Book1.xlsx";
            //connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            ////string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
            ////        "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            //using (OleDbConnection cn = new OleDbConnection(connectionString))
            //{
            //    cn.Open();
            //    OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
            //         "([Column1],[Column2],[Column3],[Column4]) " +
            //         "VALUES(@value1, @value2, @value3, @value4)", cn);
            //    cmd1.Parameters.AddWithValue("@value1", "Key1");
            //    cmd1.Parameters.AddWithValue("@value2", "Sample1");
            //    cmd1.Parameters.AddWithValue("@value3", 1);
            //    cmd1.Parameters.AddWithValue("@value4", 9);
            //    cmd1.ExecuteNonQuery();
            //}

        }

        public static void AddData(string _houseNo, string _ipAddress, double _port, double _solarSerialNo, int _solarReading, int row)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            string path = CreateDirectory();
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                //oXL.Visible = true;

                //Get a new workbook.
                //oWB = oXL.Workbooks.Open("C:\\2018\\November\\A-02.xls");
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                //oSheet.Shapes.AddPicture("http://intellibot.io/img/IB/logo.png", MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 300, 45);



                oSheet.Cells[3, 2] = "HouseNo";
                oSheet.Cells[3, 3] = "IPAddress";
                var columnHeadingsRange = oSheet.Range[
                                                        oSheet.Cells[3, 2],
                                                        oSheet.Cells[3, 3]];
                columnHeadingsRange.Interior.Color = XlRgbColor.rgbSandyBrown;


                oSheet.Cells[3, 4] = "Port";
                oSheet.Cells[3, 5] = "Serial No";
                var columnHeadingsRange_1 = oSheet.Range[
                                                        oSheet.Cells[3, 4],
                                                        oSheet.Cells[3, 5]];
                columnHeadingsRange_1.Interior.Color = XlRgbColor.rgbSandyBrown;

                oSheet.Cells[3, 6] = "DailyYeild";
                oSheet.Cells[3, 7] = "TotalYeild";
                var columnHeadingsRange_2 = oSheet.Range[
                                                     oSheet.Cells[3, 6],
                                                     oSheet.Cells[3, 7]];
                columnHeadingsRange_2.Interior.Color = XlRgbColor.rgbSandyBrown;

                oSheet.get_Range("B2", "G3").Font.Bold = true;
                oSheet.get_Range("B2", "G3").HorizontalAlignment =
                  Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                oSheet.get_Range("B2").Font.Bold = true;

                //Add table headers going cell by cell.
                //oSheet.Cells[4, 1] = "Month:";
                //oSheet.Cells[5, 1] = "House ID:";
                //oSheet.Cells[6, 1] = "Current Reading :";
                //oSheet.Cells[7, 1] = "Previous Reading :";
                //oSheet.Cells[8, 1] = "Difference Reading:";
                //oSheet.Cells[9, 1] = "Rate:";
                //oSheet.Cells[10, 1] = "Amount:";
                //oSheet.Cells[11, 1] = "Monthly Maintenance";
                //oSheet.Cells[12, 1] = "Club House Charges:";
                //oSheet.Cells[13, 1] = "Internet Charges:";
                //oSheet.Cells[14, 1] = "User Handlig Charges :";
                //oSheet.Cells[15, 1] = "Sub Total ----A:";
                //oSheet.Cells[16, 1] = "";
                //oSheet.Cells[17, 1] = "Advance:";
                //oSheet.Cells[18, 1] = "Solar:";
                //oSheet.Cells[19, 1] = "Misc:";
                //oSheet.Cells[20, 1] = "Sub Total ----B:";
                //oSheet.Cells[21, 1] = "Grand Total  (A-B):";
                //oSheet.Cells[22, 1] = "Amount in Words:";

                //oSheet.get_Range("A4", "A22").Font.Bold = false;
                //oSheet.get_Range("A4", "A22").HorizontalAlignment =
                //    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                //oSheet.get_Range("A21").Font.Bold = true;

                //oSheet.get_Range("A4", "A21").ColumnWidth = 20;
                //oSheet.get_Range("B4", "B9").ColumnWidth = 40;


                //Adding the cell value
                oSheet.Cells[row, 2] = _houseNo;
                oSheet.Cells[row, 3] = _ipAddress;
                oSheet.Cells[row, 4] = _port + "kwh";
                oSheet.Cells[row, 5] = _port + "kwh";
                oSheet.Cells[row, 6] = _port + "units";
                oSheet.Cells[row, 7] = _port;
                oSheet.Cells[row, 2] = _port;

                oSheet.get_Range("B4", "B22").Font.Bold = false;
                oSheet.get_Range("B4", "B22").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignJustify;

                oXL.Visible = false;
                oXL.UserControl = false;

                oWB.SaveAs(path + _houseNo + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Save();

                oWB.Close();
                Marshal.ReleaseComObject(oWB);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static string CreateDirectory()
        {
            string root = @"C:/" + DateTime.Now.Year.ToString() + "";
            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }
            string subdir = root + "/" + DateTime.Today.ToString("MMMM") + "/";
            if (!Directory.Exists(subdir))
            {
                Directory.CreateDirectory(subdir);
            }
            return subdir;
        }

        public static int GetMSB(int[] intValue)
        {
            try
            {
                if (intValue != null && intValue.Length > 0)
                {
                    var id = intValue[3];//4655;
                    var hexid = $"{id:X}";
                    var id1 = intValue[4];//31213;
                    var hexid1 = $"{id1:X}";
                    var resulthex = hexid + hexid1;
                    int value = Convert.ToInt32(resulthex, 16);//Convert the Hex value to Integer(MSB)
                    return value;
                }
                else
                {
                    return 0;
                }

            }
            catch (Exception)
            {
                throw;
            }

        }


        public static void Test()
        {
            System.Data.OleDb.OleDbConnection oleDbConnection;
            System.Data.OleDb.OleDbCommand oleDbCommand = new OleDbCommand();
            string sql = null;
            //string con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\\Users\\Anbarasu\\Desktop\\Test.xlsx'/Furniture.mdb";
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='C:\\Users\\Anbarasu\\Desktop\\Test.xlsx';Extended Properties=\"Excel 12.0;\"";
            // connString = "Provider=Microsoft.jet.OLEDB.4.0;Data Source='C:\\Users\\Anbarasu\\Desktop\\Test.xlsx';Extended Properties=\"Excel 8.0";
            oleDbConnection = new OleDbConnection(connString);
            oleDbConnection.Open();
            oleDbCommand.Connection = oleDbConnection;
            sql = "Insert into [sheet1$] (Username,Password) values('1','ashdfjsiahdg')";//"Select * from [Sheet1$]";
            oleDbCommand.CommandText = sql;
            oleDbCommand.ExecuteNonQuery();
            //OleDbDataAdapter da = new OleDbDataAdapter(oleDbCommand);
            //System.Data.DataTable dataTable = new System.Data.DataTable();
            //da.Fill(dataTable);
            oleDbConnection.Close();

        }

        public static void Testy()
        {
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='C:\\Users\\Anbarasu\\Desktop\\Test.xlsx';Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            OleDbConnection conn = new OleDbConnection();
            conn = new OleDbConnection(connString);
            //.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\kenny\Documents\Visual Studio 2010\Projects\Copy Cegees\Cegees\Cegees\Login.accdb";

            String Username = "Sai";//TEXTNewUser.Text;
            String Password = "baba"; //TEXTNewPass.Text;

            string sql = "UPDATE [Sheet1$] " + "SET [Username]=" + Username.Trim() + " WHERE [Password]=" + Password.Trim();

            OleDbCommand cmd = new OleDbCommand(sql);//("Insert into [sheet1$] (Username, [Password]) values(@Username, @Password)");//("INSERT into Login (Username, Password) Values(@Username, @Password)");
            cmd.Connection = conn;

            conn.Open();

            if (conn.State == ConnectionState.Open)
            {
                cmd.Parameters.Add("@Username", OleDbType.VarChar).Value = Username;
                cmd.Parameters.Add("@Password", OleDbType.VarChar).Value = Password;

                try
                {
                    cmd.ExecuteNonQuery();
                    //MessageBox.Show("Data Added");
                    conn.Close();
                }
                catch (OleDbException ex)
                {
                    //MessageBox.Show(ex.Source);
                    conn.Close();
                }
            }
            else
            {
                //MessageBox.Show("Connection Failed");
            }
        }


        public static void ReadExcel()
        {
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='C:\\Users\\Anbarasu\\Desktop\\Test.xlsx';Extended Properties=\"Excel 12.0;\"";
            string sql = "Insert into [sheet1$] (ID, Name) values('1','ashdfjsiahdg')";
            OleDbConnection oleDbConnection = new OleDbConnection();
            oleDbConnection = new OleDbConnection(connString);
            oleDbConnection.Open();
            System.Data.OleDb.OleDbCommand oleDbCommand = new OleDbCommand();
            oleDbCommand.Connection = oleDbConnection;
            oleDbCommand.CommandText = sql;
            oleDbCommand.ExecuteNonQuery();
            oleDbConnection.Close();

            //using (OleDbConnection myCon = new OleDbConnection(ConfigurationManager.ConnectionStrings["DbConn"].ToString()))
            //{..

            //string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='C:\\Users\\Anbarasu\\Desktop\\Test.xlsx';Extended Properties=\"Excel 12.0;\"";
            //OleDbConnection conn = new OleDbConnection(connString);
            //Solar solar = new Solar();
            //solar.ID = 1;
            //solar.Name = "Sai";
            //OleDbCommand cmd = new OleDbCommand();
            //cmd.CommandType = CommandType.Text;
            //cmd.CommandText = "INSERT INTO [Sheet1$] " +
            //                 "([ID], [Name]) " +
            //                 "VALUES (?, ?)";
            //cmd.Parameters.AddWithValue("@ID", solar.ID.ToString());
            //cmd.Parameters.AddWithValue("@Name", solar.Name.ToString());
            //cmd.Connection = conn;
            //conn.Open();
            //cmd.ExecuteNonQuery();

            //System.Windows.Forms.MessageBox.Show("An Item has been successfully added", "Caption", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            //}

        }

    }
}
