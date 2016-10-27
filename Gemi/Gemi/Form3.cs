using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        SqlConnection baglan = new SqlConnection("Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
        public void Dao(string veriler)
        {
            SqlDataAdapter da = new SqlDataAdapter(veriler, baglan);            //SQL BAĞLANTISI KURULDU
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }
        private void btngoruntule_Click(object sender, EventArgs e)
        {
            Dao("select distinct tip_kod as 'Gemi Tip No' , liman_no as 'Liman No' from gemitipliman");
        }
        public void combo1()
        {
            baglan.Open();
            comboBox2.Items.Clear();

            SqlCommand komut1 = new SqlCommand("select tip_kod from gemitip", baglan);          //COMBOBOX'DAN VERİLER ÇEKİLDİ
            SqlDataReader dr = komut1.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr[0]);

            }
            baglan.Close();

        }

        public void combo2()
        {
            baglan.Open();
            comboBox1.Items.Clear();

            SqlCommand komut1 = new SqlCommand("select liman_no from liman", baglan);           //COMBOBOXDAN VERİLER ÇEKİLDİ
            SqlDataReader dr = komut1.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr[0]);

            }
            baglan.Close();
        }
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form3_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();                                                             //KULLANCI YETKİLENDİRME
                button1.Hide();
                button2.Hide();
            }
            combo1();
            combo2();
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("insert into gemitipliman (tip_kod,liman_no) values (@tip_kod,@liman_no)", baglan);

                komut.Parameters.AddWithValue("@tip_kod", comboBox2.SelectedItem.ToString());           //EKLE BUTONU
                komut.Parameters.AddWithValue("@liman_no", comboBox1.SelectedItem.ToString());
                komut.ExecuteNonQuery();
                Dao("select distinct tip_kod as 'Gemi Tip No' , liman_no as 'Liman No' from gemitipliman");
                baglan.Close();


                comboBox1.Text = "";
                comboBox2.Text = "";
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int secili_alan = dataGridView1.SelectedCells[0].RowIndex;
                string tip_kod = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();
                string liman_no = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();                //DATAGRİD CELLCLİCK
                comboBox1.Text = liman_no;
                comboBox2.Text = tip_kod;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("delete gemitipliman where liman_no=@liman_no and tip_kod=@tip_kod", baglan);         //SİL BUTONU
                komut.Parameters.AddWithValue("@liman_no", dataGridView1.CurrentRow.Cells[1].Value.ToString());
                komut.Parameters.AddWithValue("@tip_kod", dataGridView1.CurrentRow.Cells[0].Value.ToString());
                komut.ExecuteNonQuery();
                Dao("select distinct tip_kod as 'Gemi Tip No' , liman_no as 'Liman No' from gemitipliman");

                baglan.Close();

                comboBox1.Text = "";
                comboBox2.Text = "";
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("update gemitipliman  set liman_no='" + comboBox1.Text + "' where tip_kod='" + comboBox2.Text + "' ", baglan);
                komut.ExecuteNonQuery();
                Dao("select distinct tip_kod as 'Gemi Tip No' , liman_no as 'Liman No' from gemitipliman");  //GÜNCELLE BUTONU
                baglan.Close();
                comboBox1.Text = "";
                comboBox2.Text = "";
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excelver = new Microsoft.Office.Interop.Excel.Application();
                Excelver.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook objBook = Excelver.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objBook.Worksheets.get_Item(1);
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range objRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[1, i + 1];         //EXCEL'E AKTAR
                    objRange.Value2 = dataGridView1.Columns[i].HeaderText;

                }
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Rows.Count; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range objRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[j + 2, i + 1];
                        objRange.Value2 = dataGridView1[i, j].Value;
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click_1(object sender, EventArgs e)
        {

            SqlDataAdapter da = new SqlDataAdapter("select distinct tip_kod as 'Gemi Tip No' , liman_no as 'Liman No' from gemitipliman where tip_kod like '%" + textBox1.Text + "%'", baglan);
            DataSet ds = new DataSet();
            try
            {
                baglan.Open();
                da.Fill(ds, "gemitipliman");
                dataGridView1.DataSource = ds.Tables["gemitipliman"];
                textBox1.Clear();


            }
            catch (Exception error)
            {

                MessageBox.Show("Hata oluştu  : " + error.Message);
            }
            finally
            {
                baglan.Close();
            }
        }

    }
}