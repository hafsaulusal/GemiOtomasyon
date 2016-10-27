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
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }
        SqlConnection baglan = new SqlConnection("Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
        public void Dao(string veriler)
        {
            SqlDataAdapter da = new SqlDataAdapter(veriler, baglan);
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void btngoruntule_Click(object sender, EventArgs e)
        {
            Dao("Select tip_kod as 'Gemi Kod',uretici as 'Gemi Üretici',max_tonaj as 'Gemi Kapasite' from gemitip ");
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("insert into gemitip (tip_kod,uretici,max_tonaj) values(@tip_kod,@uretici,@max_tonaj)", baglan);
                komut.Parameters.AddWithValue("@tip_kod", textBox1.Text);
                komut.Parameters.AddWithValue("@uretici", textBox2.Text);
                komut.Parameters.AddWithValue("@max_tonaj", textBox4.Text);
                komut.ExecuteNonQuery();
                Dao("Select tip_kod as 'Gemi Kod',uretici as 'Gemi Üretici',max_tonaj as 'Gemi Kapasite' from gemitip ");
                baglan.Close();
                textBox1.Clear();
                textBox2.Clear();
                textBox4.Clear();
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
                SqlCommand komut = new SqlCommand("delete from gemitip where tip_kod=@tip_kod", baglan);
                komut.Parameters.AddWithValue("@tip_kod", dataGridView1.CurrentRow.Cells[0].Value.ToString());
                komut.ExecuteNonQuery();
                Dao("Select tip_kod as 'Gemi Kod',uretici as 'Gemi Üretici',max_tonaj as 'Gemi Kapasite' from gemitip ");
                baglan.Close();
                textBox1.Clear();
                textBox2.Clear();
                textBox4.Clear();
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
                string uretici = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();
                string max_tonaj = dataGridView1.Rows[secili_alan].Cells[2].Value.ToString();

                textBox1.Text = tip_kod;
                textBox2.Text = uretici;
                textBox4.Text = max_tonaj;
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
                SqlCommand komut = new SqlCommand("update gemitip set uretici='" + textBox2.Text + "',max_tonaj='" + textBox4.Text + "' where tip_kod='" + textBox1.Text + "'", baglan);
                komut.ExecuteNonQuery();
                Dao("Select tip_kod as 'Gemi Kod',uretici as 'Gemi Üretici',max_tonaj as 'Gemi Kapasite' from gemitip ");
                baglan.Close();

                textBox1.Clear();
                textBox2.Clear();
                textBox4.Clear();
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
         

        }

        private void Form5_Shown(object sender, EventArgs e)
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
                    Microsoft.Office.Interop.Excel.Range objRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[1, i + 1];
                    objRange.Value2 = dataGridView1.Columns[i].HeaderText;

                }                                                           //EXCEL'E AKTAR
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
        private void Form5_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();                                     //KULLANICI YETKİLENDİRME
                button1.Hide();
                btnguncelle.Hide();
            } 
        }

        private void button4_Click(object sender, EventArgs e)
        {

            SqlDataAdapter da = new SqlDataAdapter("Select tip_kod as 'Gemi Kod',uretici as 'Gemi Üretici',max_tonaj as 'Gemi Kapasite' from gemitip where tip_kod like '%" + textBox3.Text + "%'", baglan);
            DataSet ds = new DataSet();
            try
            {
                baglan.Open();
                da.Fill(ds, "gemitip");
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
