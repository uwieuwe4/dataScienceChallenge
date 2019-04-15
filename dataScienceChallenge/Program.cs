using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Collections;
using System.IO;

namespace dataScienceChallenge
{
    class Program
    {


        public static Hashtable hashtable;
        public static Hashtable hashtableSum;
        public static List<String> users;

        static void Main(string[] args)
        {
            string filePath = "C:\\Users\\730N\\Desktop\\dataScience\\progTest.xlsx";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel._Workbook workBook = null;
            Excel._Worksheet workSheet = null;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            hashtable = new Hashtable();
            hashtableSum = new Hashtable();
            int linecount = 275840;
            users = new List<string>();
            //int linecount = 275840;

            string user;
            char gender;

            try
            {
                //Open worksheet in excel
                workBook = xlApp.ActiveWorkbook;
                workSheet = xlApp.ActiveSheet as Excel._Worksheet;
                xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                //xlApp.Visible = true;
                Console.WriteLine("Beginne Zählung");
                //iterate over cells and cound f and m 
                for (int i = 1; i < linecount; i++)
                {
                    user = xlWorkSheet.Cells[i, 1].Value.ToString();
                    gender = Convert.ToChar(xlWorkSheet.Cells[i, 5].Value());
                    adjustValues(user, gender);
                }
                Console.WriteLine("Zählung abgeschlossen. Starte Vorhersage ...");
                //make predictions based on counting
                using (StreamWriter writer = new StreamWriter("users.txt"))
                {
                    foreach (string username in users)
                    {
                        writer.WriteLine(username + "," + (int)hashtable[username] + "," + (int)hashtableSum[username]);
                    }
                }
               
                //int count;
                //for (int i = 1; i < linecount; i++)
                //{

                //    using (StreamWriter writer = new StreamWriter("important.txt"))
                //    {
                //        user = xlWorkSheet.Cells[i, 1].Value.ToString();
                //        count = getCount(user);
                //        if (count > 0)
                //        {
                //            writer.WriteLine("f,"+count);
                //            //xlWorkSheet.Cells[i, 7].Value = "f";
                //            //xlWorkSheet.Cells[i, 8].Value = count;
                //        }
                //        else if (count < 0)
                //        {
                //            writer.WriteLine("m," + count);
                //            //xlWorkSheet.Cells[i, 7].Value = "m";
                //            //xlWorkSheet.Cells[i, 8].Value = count;
                //        }
                //        else
                //        {

                //            //xlWorkSheet.Cells[i, 7].Value = "x";
                //            //xlWorkSheet.Cells[i, 8].Value = count;
                //        }
                //    }
                //}
                Console.WriteLine("Vorhersage abgeschlossen.");
                //xlApp.Visible = false;
                xlApp.UserControl = false;
                xlWorkBook.SaveAs("C:\\Users\\730N\\Desktop\\dataScience\\progTestNeu.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                xlWorkBook.Close();
            }
            catch(Exception ex)
            { Console.WriteLine(ex.Message); }
            Console.ReadKey();
            

        }

        private static int getCount(string user)
        {
            if (hashtable.ContainsKey(user))
            {
                int count = (int)hashtable[user];
                return count;
            }
            else
            {
                throw new Exception("Fehler ist passiert, für user " + user + " existieren keine Daten");
            }
        }
        private static void adjustValues(string user, char gender)
        {
            int count, countSum;
            //wenn user schon in tabelle enthalten:
            if (hashtable.ContainsKey(user))
            {
                count = (int)hashtable[user];
                countSum = (int)hashtableSum[user];
                if (gender == 'f')
                {
                    count = count+=1;
                    countSum = countSum+=1;
                    hashtable[user] = count;
                    hashtableSum[user] = countSum;
                }
            else if (gender == 'm')
                {
                    count = count -= 1;
                    countSum = countSum += 1;
                    hashtable[user] = count;
                    hashtableSum[user] = countSum;
                }
                
            else
                throw new Exception("FEHLER. WEDER F noch M");
            }
            else
            {
                if (gender == 'f')
                    hashtable.Add(user, 1);
                else if (gender == 'm')
                    hashtable.Add(user, -1);
                else
                    throw new Exception("FEHLER. WEDER F noch M");
                users.Add(user);
                hashtableSum.Add(user, 1);
            }
        }
    }
}
