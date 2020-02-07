using System;
using System.Collections.Generic;
using System.Text;
using IronXL;


namespace POMHumanity
{
    public class ExcelUtility
    {
        private static WorkBook wb = null;
        private static WorkSheet ws = null;
        public static bool OpenFile(string zaposleni)
        {
            try
            {
                if (wb == null)
                {
                    wb = WorkBook.Load(zaposleni);
                }
                else
                {
                    Console.WriteLine("Some file has already been uploaded!");
                }
                return true;
            }catch(Exception e)
            {
                Console.WriteLine(e.ToString());
                wb = null;
                return false;
            }
        }
        public static bool CloseFile()
        {
            try{
                if (wb != null)
                {
                    wb.Close();
                    wb = null;
                }
                return true;
            }catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.WriteLine("Error closing");
                return false;
            }
        }
        public static bool LoadWorkSheet(int index)
        {
            try
            {
                if (wb != null)
                {
                    ws = wb.WorkSheets[index];
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.WriteLine("Not loaded WorkSheet!");
                ws = null;
                return false;
            }
        }
        public static string GetDataAt(int row, int column)
        {
            if (ws == null)
            {
                Console.WriteLine("WorkSheet not loaded!");
                return "ERROR";
            }
            if (row < ws.Rows.Count)
            {
                RangeRow rangeRow = ws.Rows[row];
                if (column < rangeRow.Columns.Count)
                {
                    RangeColumn rangeColumn = rangeRow.Columns[column];
                    if (rangeColumn != null)
                    {
                        return rangeColumn.StringValue;
                    }
                    else
                    {
                        Console.WriteLine("There is no cell!");
                        return "ERROR";
                    }
                }
                else
                {
                    Console.WriteLine("There is no column!");
                    return "ERROR";
                }
            }
            else
            {
                Console.WriteLine("There is no row!");
                return "ERROR";
            }
        }
        //Upisuje u excel tabelu podatak
        public static bool SetData(int row, string data)
        {
            if (ws == null)
            {
                Console.WriteLine("Error setting data");
                return false;
            }

            ws.SetCellValue(row, ws.Rows[row].Columns.Count, data);
            wb.SaveAs(Constants.Excel_Path);
            return true;
        }
        //Vraca broj redova
        public static int GetRowCount()
        {
            if (ws == null)
            {
                Console.WriteLine("WorkSheet not loaded!");
                return 0;
            }
            else
            {
                return ws.Rows.Count;
            }
        }






    }
}
