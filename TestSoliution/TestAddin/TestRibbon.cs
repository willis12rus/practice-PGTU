using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace TestAddin
{
    
    public partial class TestRibbon 
    {
        
        private void TestRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            InputForm form = new InputForm();
            form.Show();

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //currentWorksheet.Cells[1, 13] = Hotel.name;
            FillingBigTable(ref currentWorksheet);
            currentWorksheet.Cells.AutoFit();
        }

        public void FillingBigTable(ref Worksheet worksheet)
        {
            List<string> categories = new List<string>();
            List<int> roomsCount = new List<int>();
            ReadData(ref worksheet, ref categories, ref roomsCount);
            //worksheet.Cells[12, 6].Value2 = worksheet.Cells[2, 1].Value2;
            /*for(int i = Hotel.currentStr + 1;i <= 13; ++i)
            {
                worksheet.Cells[i, 1].Value2 = Hotel.name;
            }*/
            for(int i = 0; i < Hotel.categoryCount; ++i)
            {
                int temp = 0;
                while(temp != roomsCount[i])
                {
                    worksheet.Cells[Hotel.currentStr+1, 1].Value2 = Hotel.name;
                    worksheet.Cells[Hotel.currentStr+1, 6].Value2 = categories[i];
                    worksheet.Cells[Hotel.currentStr+1, 7].Value2 = Hotel.entry.ToShortDateString();
                    worksheet.Cells[Hotel.currentStr+1, 9].Value2 = Hotel.departure.ToShortDateString();
                    Hotel.currentStr += 1;
                    ++temp;
                }
                temp = 0;
            }
            /*for(int i = Hotel.currentStr+1; i < roomsCount[Hotel.categoryCount]; ++i)
            {
                worksheet.Cells[i, 1] = Hotel.name;
                for(int j = 0; j < Hotel.categoryCount;++j)
                {
                    worksheet.Cells[i, 1] = Hotel.name;
                    worksheet.Cells[i, 6] = categories[i];
                    worksheet.Cells[i, 7] = Hotel.entry.ToShortDateString();
                    worksheet.Cells[i, 9] = Hotel.departure.ToShortDateString();
                }
            }*/
        }

        public void ReadData(ref Worksheet worksheet, ref List<string> categories, ref List<int> roomsCount)
        {
            int temp = 0;
            for(int i = 1; i <= Hotel.categoryCount; ++i)
            {
                categories.Add(worksheet.Cells[i+1, 1].Value2);
                roomsCount.Add(Convert.ToInt32(worksheet.Cells[i+1, 7].Value2));
                temp += roomsCount[i - 1];
            }
            roomsCount.Add(temp);
            //worksheet.Cells[12, 12].Value2 = categories[0];
        }
    }
}

public class Hotel
{
    public static string name;
    public static int categoryCount;
    public static DateTime entry;
    public static DateTime departure;
    public static int currentStr;
};
