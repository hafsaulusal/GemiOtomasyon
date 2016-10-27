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
    public partial class Form6 : Form
    {
        public Form6()
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
            //GÖRÜNTÜLE BUTONU
            Dao("set dateformat dMy;select tc as 'Personel Tc',ad_soyad as 'Personel Ad-Soyad',convert(varchar,dogum_tarih,103) as 'Personel Doğum T.',yeterlilik as 'Personel Derece' from personel ");
        }
        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("set dateformat dMy;insert into personel (tc,ad_soyad,dogum_tarih,yeterlilik) values (@tc,@ad_soyad,convert(varchar,@dogum_tarih,103),@yeterlilik)", baglan);
                komut.Parameters.AddWithValue("@tc", textBox1.Text);
                komut.Parameters.AddWithValue("@ad_soyad", textBox3.Text);
                komut.Parameters.AddWithValue("@dogum_tarih", dateTimePicker1.Text);
                komut.Parameters.AddWithValue("@yeterlilik", textBox2.Text);            //EKLE BUTONU

                komut.ExecuteNonQuery();
                Dao("set dateformat dMy;select tc as 'Personel Tc',ad_soyad as 'Personel Ad-Soyad',convert(varchar,dogum_tarih,103) as 'Personel Doğum T.',yeterlilik as 'Personel Derece' from personel ");
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();

                dateTimePicker1.Value = DateTime.Now;
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
                SqlCommand komut = new SqlCommand("delete from personel where tc=@tc", baglan);
                komut.Parameters.AddWithValue("@tc", dataGridView1.CurrentRow.Cells[0].Value.ToString());       //SİL BUTONU
                komut.ExecuteNonQuery();
                Dao("set dateformat dMy;select tc as 'Personel Tc',ad_soyad as 'Personel Ad-Soyad',convert(varchar,dogum_tarih,103) as 'Personel Doğum T.',yeterlilik as 'Personel Derece' from personel ");
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();

                dateTimePicker1.Value = DateTime.Now;
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
                baglan.Open();                                                                      //GÜNCELLE BUTONU
                SqlCommand komut = new SqlCommand("set dateformat dMy;update personel set ad_soyad='" + textBox3.Text + "',dogum_tarih='" + dateTimePicker1.Value.ToString() + "',yeterlilik='" + textBox2.Text + "'  where tc='" + textBox1.Text + "'  ", baglan);
                komut.ExecuteNonQuery();
                 Dao("set dateformat dMy;select tc as 'Personel Tc',ad_soyad as 'Personel Ad-Soyad',convert(varchar,dogum_tarih,103) as 'Personel Doğum T.',yeterlilik as 'Personel Derece' from personel ");
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();

                dateTimePicker1.Value = DateTime.Now;

            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
         
            

        }
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form6_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();                 //KULLANICI YETKİLENDİRME
                button1.Hide();
                btnguncelle.Hide();
            } 
        }

        private void Form6_Shown(object sender, EventArgs e)
        {
            textBox1.Focus();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excelver = new Microsoft.Office.Interop.Excel.Application();
                Excelver.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook objBook = Excelver.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objBook.Worksheets.get_Item(1);
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range objRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[1, i + 1]; //EXCEL'E AKTAR
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int secili_alan = dataGridView1.SelectedCells[0].RowIndex;
                string tc = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();
                string ad_soyad = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();
                string dogum_tarih = dataGridView1.Rows[secili_alan].Cells[2].Value.ToString();         //DATAGRİD CELLCLİCK
                string yeterlilik = dataGridView1.Rows[secili_alan].Cells[3].Value.ToString();



                textBox1.Text = tc;
                textBox3.Text = ad_soyad;
                dateTimePicker1.Text = dogum_tarih;
                textBox2.Text = yeterlilik;
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
          
        }

        private void button4_Click(object sender, EventArgs e)
        {

            SqlDataAdapter da = new SqlDataAdapter("set dateformat dMy;select tc as 'Personel Tc',ad_soyad as 'Personel Ad-Soyad',convert(varchar,dogum_tarih,103) as 'Personel Doğum T.',yeterlilik as 'Personel Derece' from personel where tc like '%" + textBox4.Text + "%'", baglan);
            DataSet ds = new DataSet();
            try
            {
                baglan.Open();
                da.Fill(ds, "personel");
                dataGridView1.DataSource = ds.Tables["personel"];
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
