using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public partial class Form14 : Form
    {

        public Form14()
        {
            InitializeComponent();
        }
        OleDbConnection baglan = new OleDbConnection("Provider=SQLOLEDB;Data Source=HAFSAULUSAL\\SQL;Initial Catalog=GemiDB;Persist Security Info=True;User ID=sa;password=1");
       
        private void Form14_Load(object sender, EventArgs e)
        {

            CrystalReport3 rapor = new CrystalReport3();
            CrystalDecisions.Shared.ParameterValues parametre = new
            CrystalDecisions.Shared.ParameterValues();
            CrystalDecisions.Shared.ParameterDiscreteValue deger = new
            CrystalDecisions.Shared.ParameterDiscreteValue();
            deger.Value=Form12.kodx;
            parametre.Add(deger);
            rapor.DataDefinition.ParameterFields["tcno"].ApplyCurrentValues(parametre);
            crystalReportViewer1.ReportSource = rapor;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }
       
        private void button1_Click_1(object sender, EventArgs e)
        {
  
        }
    }
}


