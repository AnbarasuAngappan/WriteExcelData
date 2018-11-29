﻿using Microsoft.Office.Interop.Excel;
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
            string ExcelFilePath = "F:\\2018\\InvertorData1.xlsx";//"C:\\Users\\AMRORGANO\\Desktop\\SolarTemplate\\InvertorData1.xlsx"; //"C:\\Users\\AMRORGANO\\Desktop\\SolarTemplate\\InvertorData1.xlsx";//"C:\\Users\\AMRORGANO\\Desktop\\InvertorData.xlsx";
            string ext = Path.GetExtension(ExcelFilePath);//string temp = Path.GetFileName(ExcelFilePath).ToLower(); 
            if (ext.Trim() == ".xls")//Connection String to Exce o90-l Workbook
            {
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ExcelFilePath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (ext.Trim() == ".xlsx")
            {
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }
            //connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
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
            foreach (DataRow datarowItem in dataTable.Rows)
            {
                var _houseNo = datarowItem.Field<string>("houseNo");
                var _ipAddress = datarowItem.Field<string>("IpAddress");
                var _port = Convert.ToInt32(datarowItem.Field<double>("port"));
                var _deviceID = Convert.ToInt32(datarowItem.Field<double>("deviceID"));
                var _regType = Convert.ToInt32(datarowItem.Field<double>("registerType"));

                var _dayYeildStartAddress = Convert.ToInt32(datarowItem.Field<double>("dayYeildStartAddress"));
                var _dayYeildLength = Convert.ToInt32(datarowItem.Field<double>("dayYeildLength"));

                var _serialnoStartAddress = Convert.ToInt32(datarowItem.Field<double>("serialnoStartAddress"));
                var _serialnoLength = Convert.ToInt32(datarowItem.Field<double>("serialnoLength"));

                var _totalYeildStartAddress = Convert.ToInt32(datarowItem.Field<double>("totalYeildStartAddress"));
                var _totalYeildLength = Convert.ToInt32(datarowItem.Field<double>("totalYeildLength"));
                
                int[] readHoldingRegisters = ModbusReading.ReadRegisterWithDeviceIDs(_ipAddress, _port, _serialnoStartAddress, _regType, _serialnoLength, Convert.ToByte(_deviceID));
                var byteresult = GetMSB(readHoldingRegisters);
                int[] dailyYeildHoldingRegisters = ModbusReading.ReadRegisterWithDeviceIDs(_ipAddress, _port, _dayYeildStartAddress, _regType, _dayYeildLength, Convert.ToByte(_deviceID));
                int[] _totalYeildHoldingRegisters = ModbusReading.ReadRegisterWithDeviceIDs(_ipAddress, _port, _totalYeildStartAddress, _regType, _totalYeildLength, Convert.ToByte(_deviceID));
                WriteExcelSolarReading(_houseNo, _ipAddress, Convert.ToString(_port), Convert.ToString(byteresult), Convert.ToString(dailyYeildHoldingRegisters[1] * 0.001), Convert.ToString(_totalYeildHoldingRegisters[2] * 0.001), DateTime.Now);

            }

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
            string root = @"F:/" + DateTime.Now.Year.ToString() + "";
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
        }
        public static void Excel(string ID, string Name)
        {
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='C:\\Users\\AMRORGANO\\Desktop\\Test.xlsx';Extended Properties=\"Excel 12.0;\"";
            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                conn.Open();
                // DbCommand also implements IDisposable
                using (OleDbCommand cmd = conn.CreateCommand())
                {
                    // create command with placeholders
                    cmd.CommandText =
                       "INSERT INTO [sheet1$] " +
                       "([ID], [Name]) " +
                       "VALUES(@ID, @Name)";
                    // add named parameters
                    cmd.Parameters.AddRange(new OleDbParameter[]
                    {
                       new OleDbParameter("@a", ID),
                       new OleDbParameter("@b", Name)

                    });

                    // execute
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
        }

        public static void WriteExcelSolarReading(string HouseID, string IpAddress, string Port, string SerialNo, string Dayyeild, string Totalyeild, DateTime dateTimetimestamp)
        {
            try
            {
                string path = CreateDirectory();
                string pathData = path + "SolarReading.xlsx";
                if (File.Exists(pathData))
                {
                    string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathData + ";Extended Properties=\"Excel 12.0;\"";//"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='C:\\Users\\AMRORGANO\\Desktop\\SolarReading.xlsx';Extended Properties=\"Excel 12.0;\"";
                    using (OleDbConnection conn = new OleDbConnection(connString))
                    {
                        conn.Open();
                        // DbCommand also implements IDisposable
                        using (OleDbCommand cmd = conn.CreateCommand())
                        {
                            // create command with placeholders
                            cmd.CommandText =
                               "INSERT INTO [sheet1$] " +
                               "([HouseID], [IpAddress],[Port],[SerialNo],[Dayyeild],[Totalyeild],[dateTimetimestamp]) " +
                               "VALUES(@HouseID, @IpAddress, @Port, @SerialNo, @Dayyeild, @Totalyeild, @dateTimetimestamp)";
                            // add named parameters
                            cmd.Parameters.AddRange(new OleDbParameter[]
                            {
                               new OleDbParameter("@HouseID", HouseID),
                               new OleDbParameter("@IpAddress", IpAddress),
                               new OleDbParameter("@Port", Port),
                               new OleDbParameter("@SerialNo",SerialNo),
                               new OleDbParameter("@Dayyeild", Dayyeild),
                               new OleDbParameter("@Totalyeild", Totalyeild),
                               new OleDbParameter("@dateTimetimestamp", dateTimetimestamp)
                            });
                            // execute
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }       
            
        }

    }
}