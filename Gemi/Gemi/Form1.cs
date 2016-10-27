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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SqlConnection baglan = new SqlConnection(@"Data Source=HAFSAULUSAL\SQL;Initial Catalog=GemiDB;User ID=sa;password=1");
        public string yetki;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                bool kontrol = false;
                string sql="select * from yetki where k_adi='"+textBox1.Text+"' AND k_sifre='"+textBox2.Text+"'";
                SqlCommand komut = new SqlCommand(sql, baglan);

                SqlDataReader oku = komut.ExecuteReader();
                while (oku.Read()) 
                {
                    yetki = oku["yetki"].ToString();
                    kontrol = true;
                    this.Hide();
                    
                }
                baglan.Close();
                if (kontrol == true)
                {
                    Form2 yeni = new Form2();
                    yeni.ShowDialog();
                   
                }
                else 
                {
                    MessageBox.Show("Lütfen gerekli bilgileri doğru giriniz");

                }
              
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
            
        }
         
        
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            textBox1.Focus();
        }
    }
}
