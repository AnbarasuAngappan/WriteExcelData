using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;

namespace WriteExcelData
{
    class Program
    {
        static void Main(string[] args)
        {

            Program program = new Program();
            

            string connectionstring = ConfigurationManager.ConnectionStrings["designdbEntities"].ConnectionString;

            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                con.Open();
                SqlCommand sqlCommand = new SqlCommand("ElectricityBill", con);
                using (var VillaDetails = sqlCommand.ExecuteReader())
                {
                    if (VillaDetails.HasRows)
                    {
                        while (VillaDetails.Read())
                        {
                            // int a = VillaDetails.GetSqlInt16(0);
                            string _month = VillaDetails.GetString(0);
                            string _houseID = VillaDetails.GetString(1);
                            double _monthStart = VillaDetails.GetDouble(2);
                            double _monthEnd = VillaDetails.GetDouble(3);
                            double _rate = VillaDetails.GetDouble(4);
                            double _amount = VillaDetails.GetDouble(5);
                            program.getData(_month, _houseID, _monthStart, _monthEnd, _rate, _amount);
                        }
                    }
                    else
                    {
                        Console.WriteLine("No records found.");
                        Console.Read();
                    }
                }
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
                oSheet.Shapes.AddPicture("http://intellibot.io/img/IB/logo.png", MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 300, 45);

                //Add table headers going cell by cell.
                oSheet.Cells[4, 1] = "House ID";
                oSheet.Cells[5, 1] = "Current Reading :";
                oSheet.Cells[6, 1] = "Previous Reading :";
                oSheet.Cells[7, 1] = "Difference Reading";
                oSheet.Cells[8, 1] = "Rate:";
                oSheet.Cells[9, 1] = "Amount :";

                oSheet.get_Range("A4", "A9").Font.Bold = true;
                oSheet.get_Range("A4", "A9").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                //Adding the cell value
                oSheet.Cells[4, 2] = _month;
                oSheet.Cells[5, 2] = _houseID;
                oSheet.Cells[6, 2] = _monthStart;
                oSheet.Cells[7, 2] = _monthEnd;
                oSheet.Cells[8, 2] = _rate;
                oSheet.Cells[9, 2] = _amount;

                oSheet.get_Range("B1", "B3").Font.Bold = false;
                oSheet.get_Range("B1", "B3").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignJustify;

                oSheet.Cells[8, 1] = "From";
                oSheet.Cells[8, 6] = "To";
                oSheet.Cells[8, 10] = "No of days";

                oSheet.get_Range("A8", "F6").Font.Bold = true;
                oSheet.get_Range("J8").Font.Bold = true;
                oSheet.get_Range("A8", "F6").VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                oSheet.Cells[9, 1] = "05-04-2018";
                oSheet.Cells[9, 6] = "15-08-2018";

                oSheet.Cells[11, 1] = "Date";
                oSheet.Cells[11, 2] = "Time";
                oSheet.Cells[11, 3] = "Import (kwh)";
                oSheet.Cells[11, 4] = "Export (kwh)";

                oSheet.get_Range("A11", "D14").Font.Bold = true;
                oSheet.get_Range("A11", "D14").VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignJustify;

                using (System.IO.StreamReader r = new System.IO.StreamReader("C:\\Users\\Anbarasu\\Desktop\\Sample.json"))
                {
                    string jsonData = r.ReadToEnd();
                    Account item = Newtonsoft.Json.JsonConvert.DeserializeObject<Account>(jsonData);
                    if (item != null)
                    {
                        int index = 12;
                        int count = 1;
                        for (int i = 1; i <= count; i++)
                        {
                            oSheet.Cells[index, 1] = item.CreatedDate;
                            oSheet.Cells[index, 2] = item.Email;
                            oSheet.Cells[index, 3] = item.Active;
                            oSheet.Cells[index, 4] = item.Active;
                            index++;
                        }
                    }
                }

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
