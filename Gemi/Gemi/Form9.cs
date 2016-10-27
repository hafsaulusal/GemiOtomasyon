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
    public partial class Form9 : Form
    {
        public Form9()
        {
            InitializeComponent();
        }
        SqlConnection baglan = new SqlConnection("Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
        public void Dao(string veriler)
        {
            SqlDataAdapter da = new SqlDataAdapter(veriler, baglan);                //SQL BAĞLANTISI KURULDU
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];

        }
        private void btngoruntule_Click(object sender, EventArgs e)
        {
            Dao("select gorev_id as 'Görev No',gorev_ad as 'Görev Ad' from gorev");
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("insert into gorev(gorev_id,gorev_ad) values (@gorev_id,@gorev_ad)", baglan);
                komut.Parameters.AddWithValue("@gorev_id", textBox1.Text);
                komut.Parameters.AddWithValue("@gorev_ad", textBox2.Text);
                komut.ExecuteNonQuery();                                           //EKLE BUTONU
                Dao("select gorev_id as 'Görev No',gorev_ad as 'Görev Ad' from gorev");
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
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
                SqlCommand komut = new SqlCommand("delete from gorev where gorev_id=@gorev_id", baglan);
                komut.Parameters.AddWithValue("@gorev_id", dataGridView1.CurrentRow.Cells[0].Value.ToString());
                komut.ExecuteNonQuery();
                Dao("select gorev_id as 'Görev No',gorev_ad as 'Görev Ad' from gorev");             //SİL BUTONU
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
          
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {                                                                           //DATAGRİD CELLCLİCK
                int secili_alan = dataGridView1.SelectedCells[0].RowIndex;
                string gorev_id = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();
                string gorev_ad = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();


                textBox1.Text = gorev_id;
                textBox2.Text = gorev_ad;
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
                SqlCommand komut = new SqlCommand("update gorev set gorev_ad='"+textBox2.Text+"' where gorev_id='" + textBox1.Text + "' ", baglan);
                komut.ExecuteNonQuery();                                                    //GÜNCELLE BUTONU
                Dao("select gorev_id as 'Görev No',gorev_ad as 'Görev Ad' from gorev");
                baglan.Close();


                textBox1.Clear();
                textBox2.Clear();
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
           
        }

        private void Form9_Shown(object sender, EventArgs e)
        {
            textBox1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
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
                {                                                                       //EXCEL'E AKTAR
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
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form9_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();                                                        //KULLANICI YETKİLENDİRME
                btnsil.Hide();
                btnguncelle.Hide();
            } 
        }

        private void button4_Click(object sender, EventArgs e)
        {

            SqlDataAdapter da = new SqlDataAdapter("select gorev_id as 'Görev No',gorev_ad as 'Görev Ad' from gorev where gorev_id like '%" + textBox3.Text + "%'", baglan);
            DataSet ds = new DataSet();
            try
            {
                baglan.Open();
                da.Fill(ds, "gorev");
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
        }

  
    
    
    }
}
