using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }


        private void button2_Click_1(object sender, EventArgs e)
        {
            Form5 yeni = new Form5();
            yeni.Show();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Form9 yeni = new Form9();
            yeni.Show();
        }

        private void btnliman_Click_1(object sender, EventArgs e)
        {
            Form4 yeni = new Form4();
            yeni.Show();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Form8 yeni = new Form8();
            yeni.Show();
        }

        private void btnpersonel_Click_1(object sender, EventArgs e)
        {
            Form6 yeni = new Form6();
            yeni.Show();
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            Form11 yeni = new Form11();
            yeni.Show();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            Form10 yeni = new Form10();
            yeni.Show();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            Form7 yeni = new Form7();
            yeni.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Form3 yeni = new Form3();
            yeni.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Form12 yeni = new Form12();
            yeni.Show();
        }
        Form1 f1 = (Form1)Application.OpenForms["Form1"];
        private void Form2_Load(object sender, EventArgs e)
        {
            if (f1.yetki =="1") 
            {
                button4.Hide();
            } 
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form13 yeni = new Form13();
            yeni.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
           
        }

        private void button10_Click_2(object sender, EventArgs e)
        {
            Form14 yeni = new Form14();
            yeni.Show();
        }


    }
}
