using System;
using System.Windows.Forms;
using Bytescout.Spreadsheet;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        string zav_people = "";
        string calc_people = "";
        string director_people = "";
        //Путь до примера
        string examplePath = "C:\\Users\\visua\\Desktop\\normalForm1\\unform1.xls";
        //Путь до конечного файла
        string forCreatePath = "C:\\Users\\visua\\Desktop\\normalForm1\\finish2.xls";
        Spreadsheet spreadsheet = new Spreadsheet();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            form.ShowDialog();
            zav_people = form.zav;
            calc_people = form.calc;
            director_people = form.dir;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Загрузка файла
            spreadsheet.LoadFromFile(examplePath);
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("стр1");
            //Заполнение номера
            worksheet.Cell("AD14").Value = textBox1.Text;
            //Заполнение дата
            worksheet.Cell("AL14").Value = dateTimePicker1.Value.ToString("dd.MM.yyyy");
            //Заполнение организации
            worksheet.Cell("A6").Value = comboBox1.Text;
            worksheet.Cell("BC6").Value = textBox2.Text;
            //Заполнение структуры
            worksheet.Cell("A8").Value = comboBox2.Text;
            worksheet.Cell("BC7").Value = textBox3.Text;
            //Заполнение наименования
            worksheet.Cell("A10").Value = comboBox3.Text;
            worksheet.Cell("BC10").Value = textBox4.Text;
            //Заполнение нулевой таблицы
            int row = dataGridView1.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                worksheet.Cell(21 + i, 2).Value = dataGridView1[0, i].Value;
                worksheet.Cell(21 + i, 7).Value = dataGridView1[1, i].Value;
            }
            //Заполнение первой таблицы
            row = dataGridView2.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                worksheet.Cell(21 + i, 9).Value = dataGridView2[0, i].Value;
                worksheet.Cell(21 + i, 12).Value = dataGridView2[1, i].Value;
                worksheet.Cell(21 + i, 14).Value = dataGridView2[2, i].Value;
            }
            //Заполнение второй таблицы
            row = dataGridView3.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                worksheet.Cell(21 + i, 17).Value = dataGridView3[0, i].Value;
                worksheet.Cell(21 + i, 20).Value = dataGridView3[1, i].Value;
                worksheet.Cell(21 + i, 22).Value = dataGridView3[2, i].Value;
            }
            //Заполнение третьей таблицы
            row = dataGridView4.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                worksheet.Cell(21 + i, 25).Value = dataGridView4[0, i].Value;
                worksheet.Cell(21 + i, 28).Value = dataGridView4[1, i].Value;
                worksheet.Cell(21 + i, 30).Value = dataGridView4[2, i].Value;
            }
            //Заполнение четвертой таблицы
            row = dataGridView5.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                worksheet.Cell(21 + i, 33).Value = dataGridView5[0, i].Value;
                worksheet.Cell(21 + i, 36).Value = dataGridView5[1, i].Value;
                worksheet.Cell(21 + i, 39).Value = dataGridView5[2, i].Value;
            }
            //Заполнение пятой таблицы
            row = dataGridView6.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                worksheet.Cell(21 + i, 42).Value = dataGridView6[0, i].Value;
                worksheet.Cell(21 + i, 45).Value = dataGridView6[1, i].Value;
                worksheet.Cell(21 + i, 49).Value = dataGridView6[2, i].Value;
            }
            //Заполнение шестой таблицы
            row = dataGridView7.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                worksheet.Cell(21 + i, 52).Value = dataGridView7[0, i].Value;
                worksheet.Cell(21 + i, 55).Value = dataGridView7[1, i].Value;
                worksheet.Cell(21 + i, 58).Value = dataGridView7[2, i].Value;
            }
            //Заполнение заведующего
            worksheet.Cell("J38").Value = zav_people;
            worksheet.Cell("R38").Value = zav_people;
            worksheet.Cell("Z38").Value = zav_people;
            worksheet.Cell("AH38").Value = zav_people;
            worksheet.Cell("AQ38").Value = zav_people;
            worksheet.Cell("BA38").Value = zav_people;
            //Заполнение калькуляции
            worksheet.Cell("J39").Value = calc_people;
            worksheet.Cell("R39").Value = calc_people;
            worksheet.Cell("Z39").Value = calc_people;
            worksheet.Cell("AH39").Value = calc_people;
            worksheet.Cell("AQ39").Value = calc_people;
            worksheet.Cell("BA39").Value = calc_people;
            //Заполнение директора
            worksheet.Cell("J40").Value = director_people;
            worksheet.Cell("R40").Value = director_people;
            worksheet.Cell("Z40").Value = director_people;
            worksheet.Cell("AH40").Value = director_people;
            worksheet.Cell("AQ40").Value = director_people;
            worksheet.Cell("BA40").Value = director_people;
            //N1
            worksheet.Cell("O33").Value = textBox19.Text;
            worksheet.Cell("J34").Value = textBox6.Text;
            worksheet.Cell("J36").Value = textBox9.Text;
            worksheet.Cell("J37").Value = textBox12.Text;
            //N2
            worksheet.Cell("W33").Value = textBox5.Text;
            worksheet.Cell("R34").Value = textBox8.Text;
            worksheet.Cell("R36").Value = textBox10.Text;
            worksheet.Cell("R37").Value = textBox7.Text;
            //N3
            worksheet.Cell("AE33").Value = textBox11.Text;
            worksheet.Cell("Z34").Value = textBox14.Text;
            worksheet.Cell("Z36").Value = textBox15.Text;
            worksheet.Cell("Z37").Value = textBox13.Text;
            //N4
            worksheet.Cell("AN33").Value = textBox16.Text;
            worksheet.Cell("AH34").Value = textBox18.Text;
            worksheet.Cell("AH36").Value = textBox20.Text;
            worksheet.Cell("AH37").Value = textBox17.Text;
            //N5
            worksheet.Cell("AX33").Value = textBox21.Text;
            worksheet.Cell("AQ34").Value = textBox23.Text;
            worksheet.Cell("AQ36").Value = textBox24.Text;
            worksheet.Cell("AQ37").Value = textBox22.Text;
            //N6
            worksheet.Cell("BG33").Value = textBox25.Text;
            worksheet.Cell("BA34").Value = textBox27.Text;
            worksheet.Cell("BA36").Value = textBox28.Text;
            worksheet.Cell("BA37").Value = textBox26.Text;
            //Сохранение файла
            if (File.Exists(forCreatePath))
            {
                File.Delete(forCreatePath);
            }
            spreadsheet.SaveAs(forCreatePath);

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = (comboBox1.SelectedIndex + 1).ToString();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox3.Text = (comboBox2.SelectedIndex + 1).ToString();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox4.Text = (comboBox3.SelectedIndex + 1).ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView6.Rows.Clear();
            dataGridView7.Rows.Clear();
            //N1
            textBox19.Text = "";
            textBox6.Text = "";
            textBox9.Text = "";
            textBox12.Text = "";
            //N2
            textBox5.Text = "";
            textBox8.Text = "";
            textBox10.Text = "";
            textBox7.Text = "";
            //N3
            textBox11.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox13.Text = "";
            //N4
            textBox16.Text = "";
            textBox18.Text = "";
            textBox20.Text = "";
            textBox17.Text = "";
            //N5
            textBox21.Text = "";
            textBox23.Text = "";
            textBox24.Text = "";
            textBox22.Text = "";
            //N6
            textBox25.Text = "";
            textBox27.Text = "";
            textBox28.Text = "";
            textBox26.Text = "";
        }
    }
}