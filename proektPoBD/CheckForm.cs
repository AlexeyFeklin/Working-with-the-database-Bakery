using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace proektPoBD
{
    public partial class CheckForm : Form
    {
        MySqlConnection connnection = new MySqlConnection("SERVER=localhost;DATABASE=bread;UID=root;PASSWORD=1111;");

        void refreshTables()
        {
            try
            {
                MySqlDataAdapter dataApdCheck = new MySqlDataAdapter("SELECT checks.*, products.*\r\nFROM checks\r\nJOIN check_products ON checks.id_check = check_products.id_check\r\nJOIN products ON check_products.id_product = products.id_products\r\n", connnection);
                DataSet dataSetCheck = new DataSet();
                dataApdCheck.Fill(dataSetCheck);
                dataGridView1.DataSource = dataSetCheck.Tables[0];
            }
            catch
            {
                MessageBox.Show("Ошибка при загрузке таблицы - Чеки!");
            }
        }

        void exportChecks()
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15;

            ExcelApp.Cells[1, 1] = "ИД ЧЕКА";
            ExcelApp.Cells[1, 2] = "ИД Пекарни";
            ExcelApp.Cells[1, 3] = "Дата";
            ExcelApp.Cells[1, 4] = "Статус оплаты";
            ExcelApp.Cells[1, 5] = "ИД Продукта";
            ExcelApp.Cells[1, 6] = "Название продукта";
            ExcelApp.Cells[1, 7] = "Цена";




            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    ExcelApp.Cells[j + 2, i + 1] = (dataGridView1[i, j].Value).ToString();
                }
            }
            ExcelApp.Visible = true;
        }

        public CheckForm()
        {
            InitializeComponent();
        }


        private void CheckForm_Load(object sender, EventArgs e)
        {
            refreshTables();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            frm1.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            exportChecks();
        }
    }
}
