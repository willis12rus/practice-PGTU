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
        public static int currentString = 1, currentCollum = 0;
        public static int entryDateCollum = 6, departureDateCollum = 8;
        public InputForm()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string hotel = textBox1.Text;
            string categoryName = textBox2.Text;
            int categoryCount = Convert.ToInt32(textBox3.Text);
            
            Worksheet currentWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            string entryDate = dateTimePicker1.Value.ToShortDateString();
            string departureDate = dateTimePicker2.Value.ToShortDateString();

            //currentWorksheet.Cells[1, 1] = temp2.AddDays(1);
            //currentWorksheet.Cells[1, 2] = "Категория";
            PrintHead(ref currentWorksheet);
            for (int i = currentString; i < categoryCount+1; ++i)
            {
                for (int j = 0; j < 9; ++j)
                {
                    if (j == 0)
                    {
                        currentWorksheet.Cells[i + 1, j + 1] = hotel;
                    }
                    else if (j == 5) { currentWorksheet.Cells[i + 1, j + 1] = categoryName; }
                    else if (j == entryDateCollum) { currentWorksheet.Cells[i + 1, j + 1] = entryDate; }
                    else if (j == departureDateCollum) { currentWorksheet.Cells[i + 1, j + 1] = departureDate; }

                }
                if(i == categoryCount) { currentString = i; }
            }
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        public void PrintHead(ref Worksheet currentWorksheet)
        {
            currentWorksheet.Cells[1, 1] = "Объект Размещения";
            currentWorksheet.Cells[1, 2] = "Группа";
            currentWorksheet.Cells[1, 3] = "Номер комнаты";
            currentWorksheet.Cells[1, 4] = "ФИО";
            currentWorksheet.Cells[1, 5] = "Кол-во чел. в номере";
            currentWorksheet.Cells[1, 6] = "Номер";
            currentWorksheet.Cells[1, 7] = "Заезд";
            currentWorksheet.Cells[1, 9] = "Выезд";
            int tempCollum = 11;
            //DateTime entry = dateTimePicker1.Value;
            DateTime departure = dateTimePicker2.Value;
            for (DateTime temp = dateTimePicker1.Value; temp <= departure; temp = temp.AddDays(1))
            {
                //temp =  temp.AddDays(1);
                currentWorksheet.Cells[1, tempCollum] = temp.ToShortDateString();
                ++tempCollum;
            };
        }
    }
}

