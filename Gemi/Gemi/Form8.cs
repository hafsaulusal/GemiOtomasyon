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
    public partial class Form8 : Form
    {
        public Form8()
        {
            InitializeComponent();
        }

       
        SqlConnection baglan = new SqlConnection("Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
        public void Dao(string veriler)
        {                                                                                                //SQL BAĞLANTISI
            SqlDataAdapter da = new SqlDataAdapter(veriler, baglan);
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }

        
        
        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("insert into gemi (gemi_no,tip_kod) values (@gemi_no,@tip_kod)", baglan);
                komut.Parameters.AddWithValue("@gemi_no", textBox1.Text);
                komut.Parameters.AddWithValue("@tip_kod", comboBox1.SelectedItem.ToString());           //EKLE BUTONU
                komut.ExecuteNonQuery();
                Dao("select distinct gemi.gemi_no as 'IMO No' ,gemitip.tip_kod as 'Gemi Tip Kod',gemitip.uretici as 'Gemi Üretici',gemitip.max_tonaj as 'GRT' from gemitip inner join gemi on gemitip.tip_kod=gemi.tip_kod");
                baglan.Close();

                textBox1.Clear();

                comboBox1.Text = "";
            
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
           

        }

        private void btngoruntule_Click(object sender, EventArgs e)
        {          
            //GÖRÜNTÜLE BUTONU
            Dao("select distinct gemi.gemi_no as 'IMO No' ,gemitip.tip_kod as 'Gemi Tip Kod',gemitip.uretici as 'Gemi Üretici',gemitip.max_tonaj as 'GRT' from gemitip inner join gemi on gemitip.tip_kod=gemi.tip_kod");
        }
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form8_Load(object sender, EventArgs e)
        {

            if (f1.yetki == "1")
            {
                btnekle.Hide();
                btnsil.Hide();
                btnguncelle.Hide();
            } 
             baglan.Open();
            comboBox1.Items.Clear();
            SqlCommand komut1 = new SqlCommand("select tip_kod from gemitip", baglan);          //COMBOBOXDAN VERİLER ÇEKİLDİ
            SqlDataReader dr = komut1.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr[0]);
            }
            baglan.Close();
        
        }
        private void btnsil_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("delete from gemi where gemi_no=@gemi_no", baglan);
                komut.Parameters.AddWithValue("@gemi_no", dataGridView1.CurrentRow.Cells[0].Value.ToString());
                komut.ExecuteNonQuery();                                                                            //SİL BUTONU
                Dao("select distinct gemi.gemi_no as 'IMO No' ,gemitip.tip_kod as 'Gemi Tip Kod',gemitip.uretici as 'Gemi Üretici',gemitip.max_tonaj as 'GRT' from gemitip inner join gemi on gemitip.tip_kod=gemi.tip_kod");
                baglan.Close();
                textBox1.Clear();
                comboBox1.Text = "";
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
                int secili_alan = dataGridView1.SelectedCells[0].RowIndex;                                            //DATAGRİD CELLCLİCK
                string gemi_no = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();
                string tip_kod = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();
                textBox1.Text = gemi_no;
                comboBox1.Text = tip_kod;
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
                SqlCommand komut = new SqlCommand("update gemi set tip_kod='" + comboBox1.Text + "' where gemi_no='" + textBox1.Text + "' ", baglan);
                komut.ExecuteNonQuery();                                                                //GÜNCELLE BUTONU
                Dao("select distinct gemi.gemi_no as 'IMO No' ,gemitip.tip_kod as 'Gemi Tip Kod',gemitip.uretici as 'Gemi Üretici',gemitip.max_tonaj as 'GRT' from gemitip inner join gemi on gemitip.tip_kod=gemi.tip_kod");
                baglan.Close();
                textBox1.Clear();
                comboBox1.Text = "";
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
          }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {                                                                                   //EXCEL'E AKTAR
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

            SqlDataAdapter da = new SqlDataAdapter("select distinct gemi.gemi_no as 'IMO No' ,gemitip.tip_kod as 'Gemi Tip Kod',gemitip.uretici as 'Gemi Üretici',gemitip.max_tonaj as 'GRT' from gemitip inner join gemi on gemitip.tip_kod=gemi.tip_kod where gemi_no like '%" + textBox2.Text + "%'", baglan);
            DataSet ds = new DataSet();
            try
            {
                baglan.Open();
                da.Fill(ds, "gemi");
                da.Fill(ds, "gemitip");
                dataGridView1.DataSource = ds.Tables["gemi"];
                dataGridView1.DataSource = ds.Tables["gemitip"];
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
