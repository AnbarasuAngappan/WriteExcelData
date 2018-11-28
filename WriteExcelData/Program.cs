using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace WriteExcelData
{
    class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();

            #region
            //string connectionstring = ConfigurationManager.ConnectionStrings["designdbEntities"].ConnectionString;
            //using (SqlConnection con = new SqlConnection(connectionstring))
            //{
            //    con.Open();
            //    SqlCommand sqlCommand = new SqlCommand("ElectricityBill", con);
            //    using (var VillaDetails = sqlCommand.ExecuteReader())
            //    {


            //        if (VillaDetails.HasRows)
            //        {
            //            while (VillaDetails.Read())
            //            {   
            //                string _month = VillaDetails.GetString(0);
            //                string _houseID = VillaDetails.GetString(1);
            //                double _monthStart = VillaDetails.GetDouble(2);
            //                double _monthEnd = VillaDetails.GetDouble(3);
            //                double _rate = VillaDetails.GetDouble(4);
            //                double _amount = VillaDetails.GetDouble(5);
            //                program.getData(_month, _houseID, _monthStart, _monthEnd, _rate, _amount);
            //            }
            //        }
            //        else
            //        {
            //            Console.WriteLine("No records found.");
            //            Console.Read();
            //        }
            //    }
            //}
            #endregion

            string ExcelFilePath = "";
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=Excel 12.0;Persist Security Info=True";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                string queryString = "SELECT * FROM [SheetName$]";

                OleDbCommand command = new OleDbCommand(queryString, connection);

                connection.Open();
                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    var val1 = reader[0].ToString();

                    //program.getData(_month, _houseID, _monthStart, _monthEnd, _rate, _amount);
                }
                reader.Close();
            }
        }

        public void getData(string _month, string _houseID, double _monthStart, double _monthEnd, double _rate, double _amount)
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
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                //oSheet.Shapes.AddPicture("http://intellibot.io/img/IB/logo.png", MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 300, 45);



                oSheet.Cells[3, 2] = "Electricity";
                oSheet.Cells[3, 3] = "DG";    
                var columnHeadingsRange = oSheet.Range[
                                                        oSheet.Cells[3, 2],
                                                        oSheet.Cells[3, 3]];
                columnHeadingsRange.Interior.Color = XlRgbColor.rgbSandyBrown;
               

                oSheet.Cells[3, 4] = "Solar";
                oSheet.Cells[3, 5] = "Water";
                var columnHeadingsRange_1 = oSheet.Range[
                                                        oSheet.Cells[3,4],
                                                        oSheet.Cells[3,5]];
                columnHeadingsRange_1.Interior.Color = XlRgbColor.rgbSandyBrown;

                oSheet.Cells[3, 6] = "LPG";
                oSheet.Cells[3, 7] = "NILL";
                var columnHeadingsRange_2 = oSheet.Range[
                                                     oSheet.Cells[3, 6],
                                                     oSheet.Cells[3, 7]];
                columnHeadingsRange_2.Interior.Color = XlRgbColor.rgbSandyBrown;

                oSheet.get_Range("B2", "G3").Font.Bold = true;
                oSheet.get_Range("B2", "G3").HorizontalAlignment =
                  Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                oSheet.get_Range("B2").Font.Bold = true;

                //Add table headers going cell by cell.
                oSheet.Cells[4, 1] = "Month:";
                oSheet.Cells[5, 1] = "House ID:";
                oSheet.Cells[6, 1] = "Current Reading :";
                oSheet.Cells[7, 1] = "Previous Reading :";
                oSheet.Cells[8, 1] = "Difference Reading:";
                oSheet.Cells[9, 1] = "Rate:";
                oSheet.Cells[10, 1] = "Amount:";
                oSheet.Cells[11, 1] = "Monthly Maintenance";
                oSheet.Cells[12, 1] = "Club House Charges:";
                oSheet.Cells[13, 1] = "Internet Charges:";
                oSheet.Cells[14, 1] = "User Handlig Charges :";
                oSheet.Cells[15, 1] = "Sub Total ----A:";
                oSheet.Cells[16, 1] = "";
                oSheet.Cells[17, 1] = "Advance:";
                oSheet.Cells[18, 1] = "Solar:";
                oSheet.Cells[19, 1] = "Misc:";
                oSheet.Cells[20, 1] = "Sub Total ----B:";
                oSheet.Cells[21, 1] = "Grand Total  (A-B):";
                oSheet.Cells[22, 1] = "Amount in Words:";

                oSheet.get_Range("A4", "A22").Font.Bold = false;
                oSheet.get_Range("A4", "A22").HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                oSheet.get_Range("A21").Font.Bold = true;

                oSheet.get_Range("A4", "A21").ColumnWidth = 20;
                oSheet.get_Range("B4", "B9").ColumnWidth = 40;


                //Adding the cell value
                oSheet.Cells[4, 2] = _month;
                oSheet.Cells[5, 2] = _houseID;
                oSheet.Cells[6, 2] = _monthStart + "kwh";
                oSheet.Cells[7, 2] = _monthEnd + "kwh";
                oSheet.Cells[8, 2] = _rate + "units";
                oSheet.Cells[9, 2] =  8;
                oSheet.Cells[10, 2] = _amount;

                oSheet.get_Range("B4", "B22").Font.Bold = false;
                oSheet.get_Range("B4", "B22").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignJustify;

                oXL.Visible = false;
                oXL.UserControl = false;
                
                oWB.SaveAs( path + _houseID+".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oWB.Close();
                Marshal.ReleaseComObject(oWB);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public string CreateDirectory()
        {
            string root = @"D:/" + DateTime.Now.Year.ToString() + "";
            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }
            string subdir = root + "/"  + DateTime.Today.ToString("MMMM") + "/";
            if (!Directory.Exists(subdir))
            {
                Directory.CreateDirectory(subdir);
            }
            return subdir;
        }        
    }

    #region
    //static void Main(string[] args)
    //{
    //    Microsoft.Office.Interop.Excel.Application oXL;
    //    Microsoft.Office.Interop.Excel._Workbook oWB;
    //    Microsoft.Office.Interop.Excel._Worksheet oSheet;
    //    Microsoft.Office.Interop.Excel.Range oRng;
    //    object misvalue = System.Reflection.Missing.Value;
    //    try
    //    {
    //        //Start Excel and get Application object.
    //        oXL = new Microsoft.Office.Interop.Excel.Application();
    //        //oXL.Visible = true;

    //        //Get a new workbook.
    //        oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
    //        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
    //        oSheet.Shapes.AddPicture("http://intellibot.io/img/IB/logo.png", MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 300, 45);

    //        //Add table headers going cell by cell.
    //        oSheet.Cells[4, 1] = "Owner Name :";
    //        oSheet.Cells[5, 1] = "Serial No :";
    //        oSheet.Cells[6, 1] = "DCU IP No :";


    //        oSheet.get_Range("A4", "A6").Font.Bold = true;
    //        oSheet.get_Range("A4", "A6").VerticalAlignment =
    //            Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

    //        //Adding the cell value
    //        oSheet.Cells[4, 2] = "Parry";
    //        oSheet.Cells[5, 2] = "SA00111";
    //        oSheet.Cells[6, 2] = "SA00111";

    //        oSheet.get_Range("B1", "B3").Font.Bold = false;
    //        oSheet.get_Range("B1", "B3").VerticalAlignment =
    //            Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignJustify;

    //        oSheet.Cells[8, 1] = "From";
    //        oSheet.Cells[8, 6] = "To";
    //        oSheet.Cells[8, 10] = "No of days";

    //        oSheet.get_Range("A8", "F6").Font.Bold = true;
    //        oSheet.get_Range("J8").Font.Bold = true;
    //        oSheet.get_Range("A8", "F6").VerticalAlignment =
    //        Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

    //        oSheet.Cells[9, 1] = "05-04-2018";
    //        oSheet.Cells[9, 6] = "15-08-2018";

    //        oSheet.Cells[11, 1] = "Date";
    //        oSheet.Cells[11, 2] = "Time";
    //        oSheet.Cells[11, 3] = "Import (kwh)";
    //        oSheet.Cells[11, 4] = "Export (kwh)";

    //        oSheet.get_Range("A11", "D14").Font.Bold = true;
    //        oSheet.get_Range("A11", "D14").VerticalAlignment =
    //        Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignJustify;

    //        using (System.IO.StreamReader r = new System.IO.StreamReader("C:\\Users\\Anbarasu\\Desktop\\Sample.json"))
    //        {
    //            string jsonData = r.ReadToEnd();
    //            Account item = Newtonsoft.Json.JsonConvert.DeserializeObject<Account>(jsonData);
    //            if (item != null)
    //            {
    //                int index = 12;
    //                int count = 1;
    //                for (int i = 1; i <= count; i++)
    //                {
    //                    oSheet.Cells[index, 1] = item.CreatedDate;
    //                    oSheet.Cells[index, 2] = item.Email;
    //                    oSheet.Cells[index, 3] = item.Active;
    //                    oSheet.Cells[index, 4] = item.Active;
    //                    index++;
    //                }
    //            }
    //        }

    //        oXL.Visible = false;
    //        oXL.UserControl = false;
    //        oWB.SaveAs("D:\\test5052.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
    //                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
    //                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
    //        oWB.Close();
    //        Marshal.ReleaseComObject(oWB);
    //    }

    //    catch (Exception ex)
    //    {
    //        throw new Exception(ex.Message);
    //    }
    //}

    #endregion

    public class Account
    {
        public string Email { get; set; }
        public bool Active { get; set; }
        public DateTime CreatedDate { get; set; }
        public IList<string> Roles { get; set; }
    }
}
