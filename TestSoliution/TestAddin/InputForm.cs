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
        public static int entryDateCollum = 2, departureDateCollum = 4, roomCategoryCollum = 5;
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

            Worksheet currentWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            string entryDate = dateTimePicker1.Value.ToShortDateString();
            string departureDate = dateTimePicker2.Value.ToShortDateString();

            //currentWorksheet.Cells[1, 1] = temp2.AddDays(1);
            //currentWorksheet.Cells[1, 2] = "Категория";
            //PrintHead(ref currentWorksheet);
            PrintLittleHead(ref currentWorksheet);

            for (int i = Hotel.currentStr; i < Hotel.categoryCount + 1; ++i)
            {
                for (int j = 0; j < 9; ++j)
                {
                    if (j == entryDateCollum) { currentWorksheet.Cells[i + 1, j + 1] = entryDate; }
                    else if (j == departureDateCollum) { currentWorksheet.Cells[i + 1, j + 1] = departureDate; }

                }
                if (i == Hotel.categoryCount-1) { Hotel.currentStr = i+3; }
            }
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

        public void PrintLittleHead(ref Worksheet currentWorksheet)
        {
            currentWorksheet.Cells[1, 1] = "Название категории";
            currentWorksheet.Cells[1, 3] = "Заезд";
            currentWorksheet.Cells[1, 5] = "Выезд";
            int tempCollum = 7;
            PrintDates(currentWorksheet, 1, ref tempCollum);
            tempCollum++;
            currentWorksheet.Cells[1, tempCollum] = "Одноместный";
            tempCollum++;
            currentWorksheet.Cells[1, tempCollum] = "Двухместный";
        }

        public void PrintBigHead(ref Worksheet currentWorksheet)
        {
            currentWorksheet.Cells[Hotel.currentStr, 1] = "Объект Размещения";
            currentWorksheet.Cells[Hotel.currentStr, 2] = "Группа";
            currentWorksheet.Cells[Hotel.currentStr, 3] = "Номер комнаты";
            currentWorksheet.Cells[Hotel.currentStr, 4] = "ФИО";
            currentWorksheet.Cells[Hotel.currentStr, 5] = "Кол-во чел. в номере";
            currentWorksheet.Cells[Hotel.currentStr, 6] = "Номер";
            currentWorksheet.Cells[Hotel.currentStr, 7] = "Заезд";
            currentWorksheet.Cells[Hotel.currentStr, 9] = "Выезд";
            int tempCollum = 11;
            PrintDates(currentWorksheet, Hotel.currentStr, ref tempCollum);
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
    }
}

