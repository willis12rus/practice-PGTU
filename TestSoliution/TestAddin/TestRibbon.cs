using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace TestAddin
{
    
    public partial class TestRibbon 
    {
        Dictionary<int, string> collumNames = new Dictionary<int, string>
        {
            { 11, "K" },
            { 12, "L"},
            { 13, "M"},
            { 14, "N"},
            { 15, "O"},
            { 16, "P"},
            { 17, "Q"},
            { 18, "R"},
            { 19, "S"},
            { 20, "T"},
            { 21, "U"},
            { 22, "V"},
            { 23, "W"},
            { 24, "X"},
            { 25, "Y"},
            { 26, "Z"},
            { 27, "AA"},
            { 28, "AB"},
            { 29, "AC"},
        };
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
            //Workbook workbook = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            //Workbook workbook = Globals.ThisAddIn.Application.Workbooks[0] ;
            Worksheet currentWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //currentWorksheet.Cells[1, 13] = Hotel.name;
            /*if(workbook.Worksheets[2] == null)
            {
                workbook.Worksheets.Add(After: currentWorksheet);

            }*/
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
            Hotel.currentStr += 1;
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
            }
            FactBlock(ref worksheet, ref roomsCount);
            DifferenceBlock(ref worksheet);
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
                categories.Add(worksheet.Cells[i+1, 6].Value2);
                roomsCount.Add(Convert.ToInt32(worksheet.Cells[i+1, 11].Value2));
                temp += roomsCount[i - 1];
            }
            roomsCount.Add(temp);
            //worksheet.Cells[12, 12].Value2 = categories[0];
        }

        public void FactBlock(ref Worksheet worksheet, ref List<int> roomsCount)
        {
            int str = Hotel.categoryCount + 4;
            int tempStr = Hotel.categoryCount * 3 + 8;
            int category = 0;
            string collumName = "";
            DateTime departureDate = Hotel.departure;
            for (int i = str; i < str + Hotel.categoryCount; ++i)
            {
                int collum = 11;
                for (DateTime temp = Hotel.entry; temp.ToShortDateString() != departureDate.AddDays(1).ToShortDateString(); temp = temp.AddDays(1))
                {
                    if (collumNames.TryGetValue(collum, out collumName))
                    {
                        string formula = "=SUM(" + collumName + tempStr + ":" + collumName;
                        int tempInt = tempStr + roomsCount[category] - 1;
                        formula += tempInt + ")";
                        worksheet.Cells[i, collum].Formula = formula;
                    }
                    //worksheet.Cells[i, collum] = temp.ToShortDateString();
                    ++collum;
                }
                tempStr += roomsCount[category];
                ++category;
            }
        }

        public void DifferenceBlock(ref Worksheet worksheet)
        {
            int str = Hotel.categoryCount * 2 + 6;
            int tempStr = 2;//+Hotel.categoryCount + 2
            //int category = 0;
            string collumName = "";
            DateTime departureDate = Hotel.departure;
            for (int i = str; i < str + Hotel.categoryCount; ++i)
            {
                int collum = 11;
                for (DateTime temp = Hotel.entry; temp.ToShortDateString() != departureDate.AddDays(1).ToShortDateString(); temp = temp.AddDays(1))
                {
                    if (collumNames.TryGetValue(collum, out collumName))
                    {
                        string formula = "=" + collumName + tempStr + " - " + collumName + (tempStr + Hotel.categoryCount + 2);
                        //formula +=  + ")";
                        worksheet.Cells[i, collum].Formula = formula;
                    }
                    //worksheet.Cells[i, collum] = temp.ToShortDateString();
                    ++collum;
                };
                ++tempStr;
            }
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
