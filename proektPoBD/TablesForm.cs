using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Excel;


namespace proektPoBD
{
    public partial class TablesForm : Form
    {
        MySqlConnection connnection = new MySqlConnection("SERVER=localhost;DATABASE=bread;UID=root;PASSWORD=1111;");

        void refreshTables()
        {
            try
            {
                MySqlDataAdapter dataApdOvercooked = new MySqlDataAdapter("select * from overcooked", connnection);
                DataSet dataSetOvercooked = new DataSet();
                dataApdOvercooked.Fill(dataSetOvercooked);
                dataGridView1.DataSource = dataSetOvercooked.Tables[0];
            }
            catch
            {
                MessageBox.Show("Ошибка при загрузке таблицы - Директор!");
            }

            try
            {
                MySqlDataAdapter dataApdDirector = new MySqlDataAdapter("select * from director", connnection);
                DataSet dataSetDirector = new DataSet();
                dataApdDirector.Fill(dataSetDirector);
                dataGridView2.DataSource = dataSetDirector.Tables[0];
            }
            catch
            {
                MessageBox.Show(" Ошибка при загрузке таблицы - Директор!");
            }

            try
            {
                MySqlDataAdapter dataApdProducts = new MySqlDataAdapter("select * from products", connnection);
                DataSet dataSetProducts = new DataSet();
                dataApdProducts.Fill(dataSetProducts);
                dataGridView3.DataSource = dataSetProducts.Tables[0];

            }
            catch
            {
                MessageBox.Show("Ошибка при загрузке таблицы - Директор!");
            }
        }

        void addDirector()
        {
            try
            {
                string query = "insert into director(surname, patronymic, name, telephone) values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "')";
                MySqlDataAdapter dataAdapterDirector = new MySqlDataAdapter(query, connnection);
                DataSet dataSetDirector = new DataSet();
                dataAdapterDirector.Fill(dataSetDirector);
                MessageBox.Show("Директор был успешно добавлен!");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при добавлении директора!");
            }
        }

        void addOvercooked()
        {
            try
            {
                string query = "insert into overcooked(nameOvercooked, adress, telephone) values('" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "')";
                MySqlDataAdapter dataAdapterOvercooked = new MySqlDataAdapter(query, connnection);
                DataSet dataSetOvercooked = new DataSet();
                dataAdapterOvercooked.Fill(dataSetOvercooked);
                MessageBox.Show("Пекарня была успешно добавлен!");
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при добавлении пекарни!");
            }
        }

        void addProduct()
        {
            try
            {
                string query = "insert into products(title, price) values('" + textBox8.Text + "','" + textBox9.Text.Replace(',', '.') + "')";
                MySqlDataAdapter dataAdapterProducts = new MySqlDataAdapter(query, connnection);
                DataSet dataSetProducts = new DataSet();
                dataAdapterProducts.Fill(dataSetProducts);
                MessageBox.Show("Продукт был успешно добавлен!");
                textBox8.Text = "";
                textBox9.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при добавлении продукта!");
            }
            
        }

        void updateOvercooked()
        {
            try
            {
                string query = "update overcooked set nameOvercooked =  '" + textBox5.Text + "', adress = '" + textBox6.Text + "', telephone = '" + textBox7.Text + "' where id = " + dataGridView1.SelectedRows[0].Cells[0].Value + " ";
                MySqlDataAdapter dataAdapterOvercooked = new MySqlDataAdapter(query, connnection);
                DataSet dataSetOvercooked = new DataSet();
                dataAdapterOvercooked.Fill(dataSetOvercooked);
                MessageBox.Show("Пекарня была успешно обновлена!");
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
            }
            catch 
            {
                MessageBox.Show("Произошла ошибка при обновлении пекарни!");
            }
        }

        void updateProducts()
        {
            try { 
            string query = "update products set title =  '" + textBox8.Text + "', price = '" + textBox9.Text.Replace(',', '.') + "' where id_products = " + dataGridView3.SelectedRows[0].Cells[0].Value + " ";
            MySqlDataAdapter dataAdapterProducts = new MySqlDataAdapter(query, connnection);
            DataSet dataSetProducts = new DataSet();
            dataAdapterProducts.Fill(dataSetProducts);
            MessageBox.Show("Продукт был успешно обновлен!");
            textBox8.Text = "";
            textBox9.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при обновлении продукта!");
            }
        }

        void updateDirector()
        {
            try
            {
                string query = "update director set surname =  '" + textBox1.Text + "', patronymic = '" + textBox2.Text + "', name = '" + textBox3.Text + "', telephone = '" + textBox4.Text + "' where id_director = " + dataGridView2.SelectedRows[0].Cells[0].Value + " ";
                MySqlDataAdapter dataAdapterDirector = new MySqlDataAdapter(query, connnection);
                DataSet dataSetDirector = new DataSet();
                dataAdapterDirector.Fill(dataSetDirector);
                MessageBox.Show("Директор был успешно обновлен!");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
            }
             
            catch 
            {
                MessageBox.Show("Произошла ошибка при обновлении директора!");
            }
        }

        void deleteOvercooked()
        {
            try
            {
                string query = "delete from overcooked where id = " + dataGridView1.SelectedRows[0].Cells[0].Value + " ";
                MySqlDataAdapter dataAdapterOvercooked = new MySqlDataAdapter(query, connnection);
                DataSet dataSetOvercooked = new DataSet();
                dataAdapterOvercooked.Fill(dataSetOvercooked);
                MessageBox.Show("Пекарня была успешно удалена!");
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при удалении пекарни!");
            }
        }

        void deleteProducts()
        {
            try
            {
                string query = "delete from products where id_products = " + dataGridView3.SelectedRows[0].Cells[0].Value + " ";
                MySqlDataAdapter dataAdapterProducts = new MySqlDataAdapter(query, connnection);
                DataSet dataSetProducts = new DataSet();
                dataAdapterProducts.Fill(dataSetProducts);
                MessageBox.Show("Продукт был успешно удален!");
                textBox8.Text = "";
                textBox9.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при удалении продукта!");
            }
        }

        void deleteDirector()
        {
            try
            {
                string query = "delete from director where id_director = " + dataGridView2.SelectedRows[0].Cells[0].Value + " ";
                MySqlDataAdapter dataAdapterDirector = new MySqlDataAdapter(query, connnection);
                DataSet dataSetDirector = new DataSet();
                dataAdapterDirector.Fill(dataSetDirector);
                MessageBox.Show("Директор был успешно удален!");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при удалении директора!");
            }
         
        }

        void searchOvercooked()
        {
         
                string query = "select * from overcooked where nameOvercooked = '" + textBox5.Text + "' OR adress = '" + textBox6.Text + "' OR telephone = '" + textBox7.Text + "' ";
                MySqlDataAdapter dataAdapterOvercooked = new MySqlDataAdapter(query, connnection);
                DataSet dataSetOvercooked = new DataSet();
                dataAdapterOvercooked.Fill(dataSetOvercooked);
                dataGridView1.DataSource = dataSetOvercooked.Tables[0];
                MessageBox.Show("Поиск пекарни был успешно воспроизведён!");
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
       
        }

        void searchProducts()
        {
            try
            {
                string query = "select * from products where title = '" + textBox8.Text + "' OR price = " + textBox9.Text;
                MySqlDataAdapter dataAdapterProducts = new MySqlDataAdapter(query, connnection);
                DataSet dataSetProducts = new DataSet();
                dataAdapterProducts.Fill(dataSetProducts);
                dataGridView3.DataSource = dataSetProducts.Tables[0];
                MessageBox.Show("Поиск пекарни был успешно воспроизведён!");
                textBox8.Text = "";
                textBox9.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при поиске продукта!");
            }
        }

        void searchDirector()
        {
            try
            {
                string query = "select * from director where surname = '"+textBox1.Text+ "' OR patronymic = '"+textBox2.Text+ "' OR name = '"+ textBox3.Text+ "' OR telephone = '"+ textBox4.Text+"'";
                MySqlDataAdapter dataAdapterDirector = new MySqlDataAdapter(query, connnection);
                DataSet dataSetDirector = new DataSet();
                dataAdapterDirector.Fill(dataSetDirector);
                dataGridView2.DataSource = dataSetDirector.Tables[0];
                MessageBox.Show("Поиск директора был успешно воспроизведён!");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при поиске директора!");
            }
        }

        void exportOvercooked()
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15;

            ExcelApp.Cells[1, 1] = "id";
            ExcelApp.Cells[1, 2] = "Название пекарни";
            ExcelApp.Cells[1, 3] = "Адресс";
            ExcelApp.Cells[1, 4] = "Телефон";


            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    ExcelApp.Cells[j + 2, i + 1] = (dataGridView1[i, j].Value).ToString();
                }
            }
            ExcelApp.Visible = true;
        }

        void exportDitector()
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15;

            ExcelApp.Cells[1, 1] = "id";
            ExcelApp.Cells[1, 2] = "Фамилия";
            ExcelApp.Cells[1, 3] = "Отчество";
            ExcelApp.Cells[1, 4] = "Имя";
            ExcelApp.Cells[1, 5] = "Телефон";


            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView2.RowCount; j++)
                {
                    ExcelApp.Cells[j + 2, i + 1] = (dataGridView2[i, j].Value).ToString();
                }
            }
            ExcelApp.Visible = true;
        }

        void exportProducts()
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15;

            ExcelApp.Cells[1, 1] = "id";
            ExcelApp.Cells[1, 2] = "Название товара";
            ExcelApp.Cells[1, 3] = "Цена";



            for (int i = 0; i < dataGridView3.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView3.RowCount; j++)
                {
                    ExcelApp.Cells[j + 2, i + 1] = (dataGridView3[i, j].Value).ToString();
                }
            }
            ExcelApp.Visible = true;
        }

        public TablesForm()
        {
            InitializeComponent();
        }

        private void TablesForm_Load(object sender, EventArgs e)
        {
            refreshTables();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            addDirector();
            refreshTables();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            addOvercooked();
            refreshTables();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            addProduct();
            refreshTables();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            updateOvercooked();
            refreshTables();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox5.Text = Convert.ToString(dataGridView1.SelectedRows[0].Cells[1].Value);
            textBox6.Text = Convert.ToString(dataGridView1.SelectedRows[0].Cells[2].Value);
            textBox7.Text = Convert.ToString(dataGridView1.SelectedRows[0].Cells[3].Value);
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = Convert.ToString(dataGridView2.SelectedRows[0].Cells[1].Value);
            textBox2.Text = Convert.ToString(dataGridView2.SelectedRows[0].Cells[2].Value);
            textBox3.Text = Convert.ToString(dataGridView2.SelectedRows[0].Cells[3].Value);
            textBox4.Text = Convert.ToString(dataGridView2.SelectedRows[0].Cells[4].Value);
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox8.Text = Convert.ToString(dataGridView3.SelectedRows[0].Cells[1].Value);
            textBox9.Text = Convert.ToString(dataGridView3.SelectedRows[0].Cells[2].Value);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            updateDirector();
            refreshTables();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            updateProducts();
            refreshTables();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            deleteOvercooked();
            refreshTables();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            deleteDirector();
            refreshTables();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            deleteProducts();
            refreshTables();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            searchOvercooked();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            searchProducts();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            searchDirector();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            refreshTables();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            exportOvercooked();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            exportProducts();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            exportDitector();   
        }

        private void button17_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            frm1.Show();
            this.Hide();
        }
    }
}
