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
    public partial class Form12 : Form
    {
        public Form12()
        {
            InitializeComponent();
        }
        SqlConnection baglan = new SqlConnection("Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
        public void Dao(string veriler)
        {
            SqlDataAdapter da = new SqlDataAdapter(veriler, baglan);                    //SQL BAĞLANTISI KURULDU
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }
        private void btngoruntule_Click(object sender, EventArgs e)
        {
            Dao("[personelgorevpro]");                                                  //PROCEDURE KOD 
            
        }

        public void combo1()
        {
            baglan.Open();
            comboBox1.Items.Clear();

            SqlCommand komut1 = new SqlCommand("select tc from personel ", baglan);     //COMBOBOXDAN VERİLER ÇEKİLDİ
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

            SqlCommand komut1 = new SqlCommand("select sefer_ayak_no from seferayak", baglan);  //COMBOBOXDAN VERİLER ÇEKİLDİ
            SqlDataReader dr = komut1.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr[0]);

            }
            baglan.Close();
        }


        public void combo3()
        {
            try
            {
                comboBox3.Items.Clear();
                    SqlCommand komut=new SqlCommand("select gorev_id,gorev_ad from gorev",baglan);       //COMBOBOXDAN VERİLER ÇEKİLDİ
                baglan.Open();
                DataTable tbl=new DataTable();
                tbl.Load(komut.ExecuteReader());
                baglan.Close();
                comboBox3.DataSource = tbl;
                comboBox3.DisplayMember = "gorev_ad";
                comboBox3.ValueMember = "gorev_id";
          }
            catch (Exception)
            {
                
                throw;
            }
        }
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form12_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();
                button1.Hide();
                button2.Hide();                                                                        //KULLANICI YETKİLENDİRME
            } 
            combo1();
            combo2();
            combo3();
           
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("set dateformat dMy;insert into personelgorev (gorev_id,sefer_ayak_no,tc,gorev_basla,gorev_bit) values (@gorev_id,@sefer_ayak_no,@tc,@gorev_basla,@gorev_bit)", baglan);
                komut.Parameters.AddWithValue("@gorev_basla", dateTimePicker1.Text);
                komut.Parameters.AddWithValue("@gorev_bit", dateTimePicker2.Text);
                komut.Parameters.AddWithValue("@gorev_id",Convert.ToInt32(comboBox3.SelectedValue));
                komut.Parameters.AddWithValue("@sefer_ayak_no", comboBox2.SelectedItem.ToString());             //EKLE BUTONU
                komut.Parameters.AddWithValue("@tc", comboBox1.SelectedItem.ToString());
                komut.ExecuteNonQuery();
                Dao("[personelgorevpro]");
                baglan.Close();

                comboBox1.Text = "";
                comboBox2.Text = "";
                comboBox3.Text = "";
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;   
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
                SqlCommand komut = new SqlCommand("delete personelgorev where sefer_ayak_no=@sefer_ayak_no and tc=@tc and gorev_id='"+Convert.ToInt32(comboBox3.SelectedValue)+"'", baglan);
                komut.Parameters.AddWithValue("@sefer_ayak_no", dataGridView1.CurrentRow.Cells[0].Value.ToString());
                komut.Parameters.AddWithValue("@tc", dataGridView1.CurrentRow.Cells[1].Value.ToString());
                komut.Parameters.AddWithValue("@gorev_id",dataGridView1.CurrentRow.Cells[2].Value.ToString());              //SİL BUTONU
                komut.ExecuteNonQuery();
                Dao("[personelgorevpro]");

                baglan.Close();

                comboBox1.Text = "";
                comboBox2.Text = "";
                comboBox3.Text = "";
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
                string gorev_bit = dataGridView1.Rows[secili_alan].Cells[4].Value.ToString();
                string gorev_basla = dataGridView1.Rows[secili_alan].Cells[3].Value.ToString();
                string gorev_id = dataGridView1.Rows[secili_alan].Cells[2].Value.ToString();
                string sefer_ayak_no = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();               //DATAGRİD CELLCLİCK
                string tc = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();
               
                
                comboBox2.Text = sefer_ayak_no;
                comboBox1.Text = tc;
                comboBox3.Text = gorev_id;
                dateTimePicker1.Text = gorev_basla;
                dateTimePicker2.Text = gorev_bit;
            
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
      }
        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("set dateformat dMy;update personelgorev set gorev_basla='" + dateTimePicker1.Value.ToString() + "',gorev_bit='" + dateTimePicker2.Value.ToString() + "' where sefer_ayak_no='" + comboBox2.Text + "' and tc='" + comboBox1.Text + "' and gorev_id='"+Convert.ToInt32(comboBox3.SelectedValue)+"' ", baglan);
                komut.ExecuteNonQuery();
                Dao("[personelgorevpro]");
                baglan.Close();
                comboBox1.Text = "";                                                //GÜNCELLE BUTONU
                comboBox2.Text = "";
                comboBox3.Text = "";
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;  
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

        private void button3_Click_1(object sender, EventArgs e)
        {
           
        }
        public static string kodx;
        private void button3_Click_2(object sender, EventArgs e)
        {
            Form14 yeni = new Form14();
            kodx = textBox1.Text;
            yeni.ShowDialog();


            /* SqlDataAdapter da = new SqlDataAdapter("set dateformat dMy;select distinct personelgorev.sefer_ayak_no as 'Sefer Ayak No',personelgorev.tc as 'Personel Tc', gorev.gorev_ad as 'Görev Ad',personelgorev.gorev_basla as 'Görev Başlangıç Tarih',personelgorev.gorev_bit as 'Görev Bitiş Tarih'  from gorev inner join personelgorev on gorev.gorev_id=personelgorev.gorev_id where tc like '%" + textBox1.Text + "%'", baglan);
             DataSet ds = new DataSet();
             try
             {
                 baglan.Open();
                 da.Fill(ds, "personelgorev");
                 da.Fill(ds, "gorev");
                 dataGridView1.DataSource = ds.Tables["personelgorev"];
                 dataGridView1.DataSource = ds.Tables["gorev"];
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
         } */

        }
    }
}
