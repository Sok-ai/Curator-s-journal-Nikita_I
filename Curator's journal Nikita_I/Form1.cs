using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Curator_s_journal_Nikita_I
{
    public partial class Form1 : Form
    {
        string[] students = {"Авилов",
                            "Бакашвили" ,
                            "Бирюков" ,
                            "Борисов" ,
                            "Бочаров" ,
                            "Вандышев" ,
                            "Гладченко" ,
                            "Гузовский" ,
                            "Еременко" ,
                            "Ищенко" ,
                            "Московкин" ,
                            };

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            monthCalendar1.MaxDate = DateTime.Today;

            int pos = 18;

            for (int i = 0; i < students.Length; i++)
            {
                Label label = new Label();
                label.Name = "stlabel" + i;
                label.AutoSize = true;
                label.Font = new Font(label.Font.Name, 9, label.Font.Style);
                label.Text = students[i];
                label.Location = new Point(300, pos + 5);
                this.Controls.Add(label);

                CheckBox checkBox = new CheckBox();
                checkBox.Name = "stCheckBox" + i;
                checkBox.Location = new Point(380, pos);
                this.Controls.Add(checkBox);

                pos += 28;
            }
            comboBox1.Items.AddRange(students);
            comboBox2.Items.AddRange(students);
        }
        List<List<string>> data = new List<List<string>>();
        private void button1_Click(object sender, EventArgs e)
        {
            bool dateExists = false;
            string datePatt = @"dd.MM.yyyy";

            foreach (var cell in data)
            {
                if (cell[0] == monthCalendar1.SelectionStart.ToString(datePatt))
                {
                    dateExists = MessageBox.Show("Такая запись уже существует, добавить ещё одну?", "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes ? false : true;
                }

            }
            if (!dateExists)
            {
                List<string> addedData = new List<string>();
                addedData.Add(monthCalendar1.SelectionStart.ToString(datePatt));
                for (int i = 0; i < students.Length; i++)
                {
                    CheckBox checkBox = (CheckBox)this.Controls["stCheckBox" + i];
                    addedData.Add((checkBox.Checked) ? "Н/б" : " ");
                 
                }
                addedData[comboBox1.SelectedIndex + 1] = (addedData[comboBox1.SelectedIndex + 1] == " ") ? "Д" : addedData[comboBox1.SelectedIndex + 1] + ", Д";
                addedData[comboBox2.SelectedIndex + 1] = (addedData[comboBox2.SelectedIndex + 1] == " ") ? "Д" : addedData[comboBox2.SelectedIndex + 1] + ", Д";

                data.Add(addedData);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var applicatoin = new Excel.Application();
            applicatoin.SheetsInNewWorkbook = 1;
            Excel.Workbook workbook = applicatoin.Workbooks.Add(System.Type.Missing);
            Excel.Worksheet worksheet = applicatoin.Worksheets.Item[1];
            worksheet.Name = "Посещаемость";

            for (int k = 0; k < students.Length; k++)
            {
                worksheet.Cells[1][k + 2] = students[k];
            }

            int i = 2;
            foreach (var cell in data)
            {
                int j = 1;
                foreach (var row in cell)
                {
                    worksheet.Cells[i][j] = row;
                    j++;
                }
                i++;
            }
            applicatoin.Visible = true;
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            string datePatt = @"dd.MM.yyyy";
            bool dateExists = false;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;

            foreach (var cell in data)
            {
                if (cell[0] == monthCalendar1.SelectionStart.ToString(datePatt))
                {
                    for (int i = 0; i < students.Length; i++)
                    {
                        CheckBox checkBox = (CheckBox)this.Controls["stCheckBox" + i];
                        checkBox.Checked = (cell[i + 1] == " " || cell[i + 1] == "Д") ? true : false;
                        if (comboBox1.SelectedIndex == -1)
                            comboBox1.SelectedIndex = (cell[i + 1] == "Д" || cell[i + 1] == "Н/б, Д") ? i : -1;
                        if (comboBox2.SelectedIndex == -1)
                            comboBox2.SelectedIndex = (cell[i + 1] == "Д" || cell[i + 1] == "Н/б, Д") ? i : -1;
                        dateExists = true;
                    }
                }
            }
            if (!dateExists)
            {
                for (int i = 0; i < students.Length; i++)
                {
                    CheckBox checkBox = (CheckBox)this.Controls["stCheckBox" + i];
                    checkBox.Checked = false;
                }
                comboBox1.SelectedIndex = -1;
                comboBox2.SelectedIndex = -1;
            }
        }
    }
}
