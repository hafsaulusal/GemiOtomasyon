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
    public partial class Form10 : Form
    {
        public Form10()
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
        {                                                                          //GÖRÜNTÜLE BUTONU
            Dao("set dateformat dMy;select sefer_no as 'Sefer No' ,convert(varchar,sefer_gun,103) as 'Sefer Tarih' from sefer");
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();                                                    //EKLE BUTONU
                SqlCommand komut = new SqlCommand("set dateformat dMy;insert into sefer (sefer_no,sefer_gun) values (@sefer_no,convert(varchar,@sefer_gun,103))", baglan);
                komut.Parameters.AddWithValue("@sefer_no", textBox1.Text);
                komut.Parameters.AddWithValue("@sefer_gun", dateTimePicker1.Text);
                komut.ExecuteNonQuery();
                Dao("set dateformat dMy;select sefer_no as 'Sefer No' ,convert(varchar,sefer_gun,103) as 'Sefer Tarih' from sefer");
                baglan.Close();

                textBox1.Clear();
                dateTimePicker1.Value = DateTime.Now;
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
                string sefer_no = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();            //DATAGRİD CELLCLİCK
                string sefer_gun = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();
                textBox1.Text = sefer_no;
                dateTimePicker1.Text = sefer_gun;
           
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
                baglan.Open();                                                                          //SİL BUTONU
                SqlCommand komut = new SqlCommand("delete from sefer where sefer_no=@sefer_no", baglan);
                komut.Parameters.AddWithValue("@sefer_no", dataGridView1.CurrentRow.Cells[0].Value.ToString());
                komut.ExecuteNonQuery();
                Dao("set dateformat dMy;select sefer_no as 'Sefer No' ,convert(varchar,sefer_gun,103) as 'Sefer Tarih' from sefer");
                baglan.Close();


                textBox1.Clear();
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
                baglan.Open();
                SqlCommand komut = new SqlCommand("set dateformat dMy;update sefer set sefer_gun='" + dateTimePicker1.Text + "' where sefer_no='" + textBox1.Text + "' ", baglan);
                komut.ExecuteNonQuery();
                Dao("set dateformat dMy;select sefer_no as 'Sefer No' ,convert(varchar,sefer_gun,103) as 'Sefer Tarih' from sefer");
                baglan.Close();                                     //GÜNCELLE BUTONU

                textBox1.Clear();
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
            {                                                      //EXCEL'E AKTAR
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
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form10_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();
                btnsil.Hide();              //KULLANCI YETKİLENDİRME
                btnguncelle.Hide();
            } 
        }

        private void button4_Click(object sender, EventArgs e)
        {

            SqlDataAdapter da = new SqlDataAdapter("set dateformat dMy;select sefer_no as 'Sefer No' ,convert(varchar,sefer_gun,103) as 'Sefer Tarih' from sefer where sefer_no like '%" + textBox2.Text + "%'", baglan);
            DataSet ds = new DataSet();
            try
            {
                baglan.Open();
                da.Fill(ds, "sefer");
                dataGridView1.DataSource = ds.Tables["sefer"];
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
