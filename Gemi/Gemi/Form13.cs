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
    public partial class Form13 : Form
    {
        public Form13()
        {
            InitializeComponent();
        }
        SqlConnection baglan = new SqlConnection("Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
        public void Dao(string veriler)
        {
            SqlDataAdapter da = new SqlDataAdapter(veriler, baglan);                        //SQL BAĞLANTISI KURULDU
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }
        private void btngoruntule_Click(object sender, EventArgs e)
        {
            Dao("[seferolaypro]");                                                          //GÖRÜNTÜLE BUTONU
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("set dateformat dMy;insert into seferolay (giris_zaman,cıkıs_zaman,sefer_ayak_no,gemi_no,liman_no) values (convert(varchar,@giris_zaman,103),convert(varchar,@cıkıs_zaman,103),@sefer_ayak_no,@gemi_no,@liman_no)", baglan);
                
                komut.Parameters.AddWithValue("@giris_zaman", dateTimePicker2.Text);
                komut.Parameters.AddWithValue("@cıkıs_zaman", dateTimePicker3.Text);
                komut.Parameters.AddWithValue("@sefer_ayak_no", comboBox1.SelectedItem.ToString());         //EKLE BUTONU
                komut.Parameters.AddWithValue("@gemi_no", comboBox2.SelectedItem.ToString());
                komut.Parameters.AddWithValue("@liman_no", comboBox3.SelectedItem.ToString());
                komut.ExecuteNonQuery();
                Dao("[seferolaypro]");
                baglan.Close();
           
                comboBox1.Text = "";
                comboBox2.Text = "";
                comboBox3.Text = "";
                dateTimePicker2.Value = DateTime.Now;
                dateTimePicker3.Value = DateTime.Now;    
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

            SqlCommand komut1 = new SqlCommand("select sefer_ayak_no from seferayak ", baglan); //COMBOBOXDAN VERİLER ÇEKİLDİ
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

            SqlCommand komut1 = new SqlCommand("select gemi_no from gemi", baglan);             //COMBOBOXDAN VERİLER ÇEKİLDİ
            SqlDataReader dr = komut1.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr[0]);

            }
            baglan.Close();
        }
         public void combo3()
        {
            baglan.Open();
            comboBox3.Items.Clear();

            SqlCommand komut1 = new SqlCommand("select liman_no from liman", baglan);           //COMBOBOXDAN VERİLER ÇEKİLDİ
            SqlDataReader dr = komut1.ExecuteReader();
            while (dr.Read())
            {
                comboBox3.Items.Add(dr[0]);

            }
            baglan.Close();
        }
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form13_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();
                button1.Hide();                                     //KULLANICI YETKİLENDİRME
                btnguncelle.Hide();

            } 
            combo1();
            combo2();
            combo3();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("delete seferolay where sefer_ayak_no=@sefer_ayak_no and gemi_no=@gemi_no and liman_no=@liman_no", baglan);
                komut.Parameters.AddWithValue("@sefer_ayak_no", dataGridView1.CurrentRow.Cells[2].Value.ToString());
                komut.Parameters.AddWithValue("@gemi_no", dataGridView1.CurrentRow.Cells[3].Value.ToString());
                komut.Parameters.AddWithValue("@liman_no", dataGridView1.CurrentRow.Cells[4].Value.ToString());
                komut.ExecuteNonQuery();
                Dao("[seferolaypro]");                              //PROCEDURE KOD

                baglan.Close();
              
                comboBox1.Text = "";
                comboBox2.Text = "";
                comboBox3.Text = "";
                dateTimePicker2.Value = DateTime.Now;
                dateTimePicker3.Value = DateTime.Now; 
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
              
                string giris_zaman = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();
                string cıkıs_zaman = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();
                string sefer_ayak_no = dataGridView1.Rows[secili_alan].Cells[2].Value.ToString();
                string gemi_no = dataGridView1.Rows[secili_alan].Cells[3].Value.ToString();
                string liman_no = dataGridView1.Rows[secili_alan].Cells[4].Value.ToString();  //DATAGRİD CELLCLİCK
                dateTimePicker2.Text = giris_zaman;
                dateTimePicker3.Text = cıkıs_zaman;
                comboBox1.Text = sefer_ayak_no;
                comboBox2.Text = gemi_no;
                comboBox3.Text = liman_no;
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
                SqlCommand komut = new SqlCommand("set dateformat dmy;update seferolay set giris_zaman='" + dateTimePicker2.Value.ToString() + "',cıkıs_zaman='" + dateTimePicker3.Value.ToString() + "' where gemi_no='" + comboBox2.Text + "'  and liman_no='" + comboBox3.Text + "' and sefer_ayak_no='" + comboBox1.Text + "'  ", baglan);
                komut.ExecuteNonQuery();
                Dao("[seferolaypro]");                                                         //GÜNCELLE BUTONU
                baglan.Close();
                comboBox1.Text = "";
                comboBox2.Text = "";
                comboBox3.Text = "";
                dateTimePicker2.Value = DateTime.Now;
                dateTimePicker3.Value = DateTime.Now;   
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excelver = new Microsoft.Office.Interop.Excel.Application();
                Excelver.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook objBook = Excelver.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objBook.Worksheets.get_Item(1);           //EXCEL'E AKTAR
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

            SqlDataAdapter da = new SqlDataAdapter("set dateformat dMy;select distinct seferayak.sefer_ayak_no as 'Sefer Ayak No',convert(varchar,seferayak.planlanan_giris,103) as 'Liman Giriş ',convert(varchar,seferayak.planlanan_cikis,103) as 'Liman Çıkış', sefer.sefer_no as 'Sefer No', convert(varchar,sefer.sefer_gun,103) as'Sefer Günü',liman.liman_no as'Liman No',liman.liman_ad as 'Liman Ad' from seferayak inner join sefer on sefer.sefer_no=seferayak.sefer_no inner join liman on liman.liman_no=seferayak.liman_no where sefer_ayak_no like '%" + textBox1.Text + "%'", baglan);
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
