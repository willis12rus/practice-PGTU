using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace TestAddin
{   

    public partial class InputForm : Form
    {
        public static int currentCollum = 0;
        public static int entryDateCollum = 6, departureDateCollum = 8, roomCategoryCollum = 5;
        public InputForm()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Hotel.name = textBox1.Text;
            Hotel.categoryCount = Convert.ToInt32(textBox2.Text);
            Hotel.entry = dateTimePicker1.Value;
            Hotel.departure = dateTimePicker2.Value;
            Hotel.currentStr = 0;

            Worksheet currentWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            currentWorksheet.Cells.Clear();
            currentWorksheet.Name = Hotel.name;
            string entryDate = dateTimePicker1.Value.ToShortDateString();
            string departureDate = dateTimePicker2.Value.ToShortDateString();

            //currentWorksheet.Cells[1, 1] = temp2.AddDays(1);
            //currentWorksheet.Cells[1, 2] = "Категория";
            //PrintHead(ref currentWorksheet);

            PrintLittleHead(ref currentWorksheet, "Блок в отеле");
            /*for (int i = Hotel.currentStr; i < Hotel.categoryCount + Hotel.currentStr; ++i)
            {
                for (int j = 3; j < 12; ++j)
                {
                    if (j == entryDateCollum && i != 0) { currentWorksheet.Cells[i + 1, j + 1] = entryDate; }
                    else if (j == departureDateCollum && i != 0) { currentWorksheet.Cells[i + 1, j + 1] = departureDate; }

                }
                if (i == Hotel.categoryCount-1) { Hotel.currentStr += Hotel.categoryCount; }
            }*/
            FillingLittleTable(ref currentWorksheet);

            
            PrintLittleHead(ref currentWorksheet, "Фактический блок");
            FillingLittleTable(ref currentWorksheet);

            //Hotel.currentStr += Hotel.categoryCount+1;
            PrintLittleHead(ref currentWorksheet, "Разница");
            FillingLittleTable(ref currentWorksheet);

            PrintBigHead(ref currentWorksheet);
            currentWorksheet.Columns.AutoFit();
            this.Close();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void InputForm_Load(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        public void PrintLittleHead(ref Worksheet currentWorksheet, string title)
        {
            currentWorksheet.Cells[Hotel.currentStr+1, 4].Value2 = title;
            int temp = Hotel.currentStr+1;
            int strPointer = 2;
            for(int i = 0; i < Hotel.categoryCount; ++i)
            {
                currentWorksheet.Cells[temp+1, 5].Value2 = i + 1;
                if(title != "Блок в отеле")
                {
                    string formula = "=F" +strPointer;
                    currentWorksheet.Cells[temp + 1, 6].Formula = formula;
                    ++strPointer;
                }
                ++temp;
            }
            currentWorksheet.Cells[Hotel.currentStr+1, 6] = "Название категории";
            currentWorksheet.Cells[Hotel.currentStr+1, 7] = "Въезд";
            currentWorksheet.Cells[Hotel.currentStr+1, 9] = "Выезд";
            int tempCollum = 11;
            PrintDates(currentWorksheet, Hotel.currentStr+1, ref tempCollum);
            tempCollum++;
            currentWorksheet.Cells[Hotel.currentStr+1, tempCollum] = "Одноместный";
            tempCollum++;
            currentWorksheet.Cells[Hotel.currentStr+1, tempCollum] = "Двухместный";
        }

        public void PrintBigHead(ref Worksheet currentWorksheet)
        {
            currentWorksheet.Cells[Hotel.currentStr+1, 1] = "Объект Размещения";
            currentWorksheet.Cells[Hotel.currentStr+1, 2] = "Группа";
            currentWorksheet.Cells[Hotel.currentStr+1, 3] = "Номер комнаты";
            currentWorksheet.Cells[Hotel.currentStr+1, 4] = "ФИО";
            currentWorksheet.Cells[Hotel.currentStr+1, 5] = "Кол-во чел. в номере";
            currentWorksheet.Cells[Hotel.currentStr+1, 6] = "Номер";
            currentWorksheet.Cells[Hotel.currentStr+1, 7] = "Въезд";
            currentWorksheet.Cells[Hotel.currentStr+1, 9] = "Выезд";
            int tempCollum = 11;
            PrintDates(currentWorksheet, Hotel.currentStr+1, ref tempCollum);
        }

        public void PrintDates(Worksheet currentWorksheet, int str, ref int collum)
        {

            DateTime entry = dateTimePicker1.Value;
            DateTime departure = dateTimePicker2.Value;
            for (DateTime temp = entry; temp.ToShortDateString() != departure.AddDays(1).ToShortDateString(); temp = temp.AddDays(1))
            {
                //temp =  temp.AddDays(1);
                currentWorksheet.Cells[str, collum] = temp.ToShortDateString();
                ++collum;
            };
        }

        public void FillingLittleTable(ref Worksheet worksheet)
        {
            for (int i = Hotel.currentStr+1; i <= Hotel.categoryCount + Hotel.currentStr; ++i)
            {
                for (int j = 3; j < 12; ++j)
                {
                    if (j == entryDateCollum && i != 0) { worksheet.Cells[i + 1, j + 1] = Hotel.entry.ToShortDateString(); }
                    else if (j == departureDateCollum && i != 0) { worksheet.Cells[i + 1, j + 1] = Hotel.departure.ToShortDateString(); }

                }
            }
            Hotel.currentStr += Hotel.categoryCount + 2;
        }
    }
}

