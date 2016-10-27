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
    public partial class Form7 : Form
    {
        public Form7()
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
            Dao("select k_adi as 'Kullanıcı Adı',k_sifre as 'Şifre',yetki as 'Kullanıcı Yetki' from yetki"); //GÖRÜNTÜLE BUTONU
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("insert into yetki (k_adi,k_sifre,yetki) values (@k_adi,@k_sifre,@yetki)", baglan);
              
                komut.Parameters.AddWithValue("@k_adi", textBox1.Text);
                komut.Parameters.AddWithValue("@k_sifre", textBox2.Text);                                   //EKLE BUTONU
                komut.Parameters.AddWithValue("@yetki", textBox3.Text);
                komut.ExecuteNonQuery();
                Dao("select k_adi as 'Kullanıcı Adı',k_sifre as 'Şifre',yetki as 'Kullanıcı Yetki' from yetki");
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
           
        }

        private void Form7_Shown(object sender, EventArgs e)
        {
            textBox1.Focus();
        }

         private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int secili_alan = dataGridView1.SelectedCells[0].RowIndex;
               
                string k_adi = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();
                string k_sifre = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();         //DATAGRİD CELLCLİCK
                string yetki = dataGridView1.Rows[secili_alan].Cells[2].Value.ToString();
                textBox1.Text = k_adi;
                textBox2.Text = k_sifre;
                textBox3.Text = yetki;
                
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
           
         }

        private void btnsil_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("delete yetki where k_adi=@k_adi", baglan);
                komut.Parameters.AddWithValue("@k_adi", dataGridView1.CurrentRow.Cells[0].Value.ToString()); //SİL BUTONU
                komut.ExecuteNonQuery();
                Dao("select k_adi as 'Kullanıcı Adı',k_sifre as 'Şifre',yetki as 'Kullanıcı Yetki' from yetki");
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                
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
                SqlCommand komut = new SqlCommand("update yetki set  k_sifre='" + textBox2.Text + "',yetki='" + textBox3.Text + "'   where k_adi='" + textBox1.Text + "'  ", baglan);
                komut.ExecuteNonQuery();
                Dao("select id as 'ID',k_adi as 'Kullanıcı Adı',k_sifre as 'Şifre',yetki as 'Kullanıcı Yetki' from yetki");         //GÜNCELLE BUTONU
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
          
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
           
        }
    }
}