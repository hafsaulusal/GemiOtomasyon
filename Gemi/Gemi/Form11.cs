using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace WindowsFormsApplication1
{
    public partial class Form11 : Form
    {
        public Form11()
        {
            InitializeComponent();
        }
        SqlConnection baglan = new SqlConnection("Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
        public void Dao(string veriler)
        {
            SqlDataAdapter da = new SqlDataAdapter(veriler, baglan);
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];                                                //SQL BAĞLANTISI KURULDU
        }
        private void btngoruntule_Click(object sender, EventArgs e)
        {
                Dao("[seferayakpro]");
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("set dateformat dMy;insert into seferayak (sefer_ayak_no,planlanan_giris,planlanan_cikis,sefer_no,liman_no) values (@sefer_ayak_no,convert(varchar,@planlanan_giris,103),convert(varchar,@planlanan_cikis,103),@sefer_no,@liman_no)", baglan);
                komut.Parameters.AddWithValue("@sefer_ayak_no", textBox1.Text);
                komut.Parameters.AddWithValue("@sefer_no", comboBox1.SelectedItem.ToString());
                komut.Parameters.AddWithValue("@liman_no", comboBox2.SelectedItem.ToString());
                komut.Parameters.AddWithValue("@planlanan_giris", dateTimePicker1.Text);
                komut.Parameters.AddWithValue("@planlanan_cikis", dateTimePicker2.Text);            //EKLE BUTONU
                komut.ExecuteNonQuery();
                Dao("[seferayakpro]");
                baglan.Close();

                textBox1.Clear();
                comboBox1.Text = "";
                comboBox2.Text = "";
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }



        }

        public void combo1()
        {
            baglan.Open();
            comboBox1.Items.Clear();

            SqlCommand komut1 = new SqlCommand("select sefer_no from sefer", baglan);           //COMBOBOXDAN VERİLER ÇEKİLDİ
            SqlDataReader dr = komut1.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr[0]);

            }
            baglan.Close();
         }
        public void combo2()
        {
            baglan.Open();
            comboBox2.Items.Clear();

            SqlCommand komut1 = new SqlCommand("select liman_no from liman", baglan);           //COMBOBOXDAN VERİLER ÇEKİLDİ   
            SqlDataReader dr = komut1.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr[0]);

            }
            baglan.Close();
        }
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form11_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();
                button1.Hide();
                btnguncelle.Hide();                                                             //KULLANICI YETKİLENDİRME
            }
            combo1();
            combo2();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("delete seferayak where sefer_ayak_no=@sefer_ayak_no", baglan);
                komut.Parameters.AddWithValue("@sefer_ayak_no", dataGridView1.CurrentRow.Cells[0].Value.ToString());
                komut.ExecuteNonQuery();
                Dao("[seferayakpro]");                                                                      //SİL BUTONU
                baglan.Close();


                textBox1.Clear();
                comboBox1.Text = "";
                comboBox2.Text = "";
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
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
                string sefer_ayak_no = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();
                string sefer_no = dataGridView1.Rows[secili_alan].Cells[3].Value.ToString();                //DATAGRİD CELLCLİCK
                string liman_no = dataGridView1.Rows[secili_alan].Cells[5].Value.ToString();
                string planlanan_giris = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();
                string planlanan_cikis = dataGridView1.Rows[secili_alan].Cells[2].Value.ToString();

                textBox1.Text = sefer_ayak_no;
                comboBox1.Text = sefer_no;
                comboBox2.Text = liman_no;
                dateTimePicker1.Text = planlanan_giris;
                dateTimePicker2.Text = planlanan_cikis;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void btnguncelle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("set dateformat dMy;update seferayak set planlanan_giris='" + dateTimePicker1.Text + "',planlanan_cikis='" + dateTimePicker2.Text + "',sefer_no='" + comboBox1.Text + "',liman_no='" + comboBox2.Text + "'  where sefer_ayak_no='" + textBox1.Text + "' ", baglan);
                komut.ExecuteNonQuery();
                Dao("[seferayakpro]");
                baglan.Close();                     //GÜNCELLEME BUTONU

                textBox1.Clear();
                comboBox1.Text = "";
                comboBox2.Text = "";
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
         }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {                                       //EXCEL'E AKTAR
                Microsoft.Office.Interop.Excel.Application Excelver = new Microsoft.Office.Interop.Excel.Application();
                Excelver.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook objBook = Excelver.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objBook.Worksheets.get_Item(1);
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range objRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[1, i + 1];
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
            
            SqlDataAdapter da = new SqlDataAdapter("set dateformat dMy;select distinct seferayak.sefer_ayak_no as 'Sefer Ayak No', convert(varchar,seferayak.planlanan_giris,103) as 'Liman Giriş ',convert(varchar,seferayak.planlanan_cikis,103) as 'Liman Çıkış', sefer.sefer_no as 'Sefer No', convert(varchar,sefer.sefer_gun,103) as'Sefer Günü',liman.liman_no as'Liman No',liman.liman_ad as 'Liman Ad' from seferayak inner join sefer on sefer.sefer_no=seferayak.sefer_no inner join liman on liman.liman_no=seferayak.liman_no where sefer_ayak_no like '%" + textBox2.Text + "%'", baglan);
            DataSet ds = new DataSet();
            try
            {
                baglan.Open();
                da.Fill(ds, "seferayak");
                da.Fill(ds, "sefer");
                da.Fill(ds, "liman");
                dataGridView1.DataSource = ds.Tables["seferayak"];
                dataGridView1.DataSource = ds.Tables["sefer"];
                dataGridView1.DataSource = ds.Tables["liman"];
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
