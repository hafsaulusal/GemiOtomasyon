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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }
        SqlConnection baglan=new SqlConnection("Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
        public void Dao(string veriler) 
        {
            SqlDataAdapter da = new SqlDataAdapter(veriler, baglan);
            DataSet ds = new DataSet ();                                        //SQL BAĞLANTISI KURULDU
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
               
        }
        
        private void btngoruntule_Click(object sender, EventArgs e)
        {
           Dao("select liman.liman_no as 'Liman No',liman.liman_ad as 'Liman Ad',limanbilgi.liman_konum as 'Liman Koordinat ', liman.max_kapasite 'Max Kapasite' from limanbilgi inner join liman on limanbilgi.id=liman.id");
            //GÖRÜNTÜLE BUTONU
        }

        private void btnekle_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("insert into liman(liman_no,liman_ad,max_kapasite,id) values(@liman_no,@liman_ad,@max_kapasite,@id)", baglan);
                komut.Parameters.AddWithValue("@id", Convert.ToInt32(comboBox1.SelectedValue));
                komut.Parameters.AddWithValue("@liman_no", textBox1.Text);
                komut.Parameters.AddWithValue("@liman_ad", textBox2.Text);
                komut.Parameters.AddWithValue("@max_kapasite", textBox3.Text);                  //EKLE BUTONU
                komut.ExecuteNonQuery();
                Dao("select liman.liman_no as 'Liman No',liman.liman_ad as 'Liman Ad',limanbilgi.liman_konum as 'Liman Koordinat ', liman.max_kapasite 'Max Kapasite'from limanbilgi inner join liman on limanbilgi.id=liman.id");

                baglan.Close();

                textBox1.Clear();
                textBox3.Clear();
                textBox2.Clear();
                comboBox1.Text = "";
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
                SqlCommand komut = new SqlCommand("delete from liman where liman_no=@liman_no", baglan);
                komut.Parameters.AddWithValue("@liman_no", dataGridView1.CurrentRow.Cells[0].Value.ToString());
                komut.ExecuteNonQuery();                //SİL BUTONU
                Dao("select liman.liman_no as 'Liman No',liman.liman_ad as 'Liman Ad',limanbilgi.liman_konum as 'Liman Koordinat ', liman.max_kapasite 'Max Kapasite'from limanbilgi inner join liman on limanbilgi.id=liman.id");
                baglan.Close();

                textBox1.Clear();
                textBox3.Clear();
                textBox2.Clear();
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
                int secili_alan = dataGridView1.SelectedCells[0].RowIndex;
                string liman_no = dataGridView1.Rows[secili_alan].Cells[0].Value.ToString();
                string max_kapasite = dataGridView1.Rows[secili_alan].Cells[3].Value.ToString();
                string liman_ad = dataGridView1.Rows[secili_alan].Cells[1].Value.ToString();                //DATAGRİD CELLCLİCK
                string id = dataGridView1.Rows[secili_alan].Cells[2].Value.ToString();
                textBox1.Text = liman_no;
                textBox2.Text = liman_ad;
                comboBox1.Text =id;
                textBox3.Text = max_kapasite;
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
                SqlCommand komut = new SqlCommand("update liman set liman_ad='" + textBox2.Text + "',id='" + Convert.ToInt32(comboBox1.SelectedValue) + "',max_kapasite='" + textBox3.Text + "' where liman_no='" + textBox1.Text + "'  ", baglan);
                komut.ExecuteNonQuery();                //GÜNCELLE BUTONU
                Dao("select liman.liman_no as 'Liman No',liman.liman_ad as 'Liman Ad',limanbilgi.liman_konum as 'Liman Koordinat ', liman.max_kapasite 'Max Kapasite'from limanbilgi inner join liman on limanbilgi.id=liman.id");
                baglan.Close();
                textBox1.Clear();
                textBox3.Clear();
                textBox2.Clear();
                comboBox1.Text = "";
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
                
                
               
        }
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form4_Load(object sender, EventArgs e)
        {
            if (f1.yetki == "1")
            {
                btnekle.Hide();
                btnsil.Hide();
                btnguncelle.Hide();
            }

            combo3();
        }

        public void combo3()
        {
            try
            {
                                        //COMBOBOX'A VERİLER ÇEKİLDİ
                SqlCommand komut = new SqlCommand("select id,liman_konum from limanbilgi", baglan);
                baglan.Open();
                comboBox1.Items.Clear();
                DataTable tbl = new DataTable();
                tbl.Load(komut.ExecuteReader());
                baglan.Close();

                comboBox1.DataSource = tbl;
                comboBox1.DisplayMember = "liman_konum";
                comboBox1.ValueMember = "id";



            }
            catch (Exception)
            {

                throw;
            }
        }
        private void Form4_Shown(object sender, EventArgs e)
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
                                                                                //EXCEL'E AKTAR
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

            SqlDataAdapter da = new SqlDataAdapter(" select liman.liman_no as 'Liman No',liman.liman_ad as 'Liman Ad',limanbilgi.liman_konum as 'Liman Koordinat ', liman.max_kapasite 'Max Kapasite' from limanbilgi inner join liman on limanbilgi.id=liman.id where liman.liman_no like '%" + textBox4.Text + "%'", baglan);
            DataSet ds = new DataSet();
            try
            {
                baglan.Open();
                da.Fill(ds, "liman");
                da.Fill(ds, "limanbilgi");
                dataGridView1.DataSource = ds.Tables["liman"];
                dataGridView1.DataSource = ds.Tables["limanbilgi"];
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
