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

namespace Access {
    public partial class Form1 : Form {

        public static String adres = null;
        //OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Erdo\Desktop\Eksper2000\yeni\dataex\EKS2000\RAPOR.MDB");
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + adres);


        public Form1() {
            InitializeComponent();

        }



        private void button1_Click(object sender, EventArgs e) {

            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + adres);
            con.Open();
            DataTable dt = con.GetSchema("Tables");
            

            con.Close(); 
            
            foreach (DataRow dataRow in dt.Rows) {

                comboBox1.Items.Add(dataRow["TABLE_NAME"].ToString().Trim());
                    
            }
        }

        private void button2_Click(object sender, EventArgs e) {

            if (openFileDialog1.ShowDialog() == DialogResult.OK) {
                adres = openFileDialog1.FileName;
            }

            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + adres);
            con.Open();
            DataTable dt = con.GetSchema("Tables");
            con.Close();

            foreach (DataRow dataRow in dt.Rows) {

                comboBox1.Items.Add(dataRow["TABLE_NAME"].ToString().Trim());

            }
        }

        private void comboBox1_SelectedValueChanged_1(object sender, EventArgs e) {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + adres);

            con.Open();

            OleDbDataAdapter da = new OleDbDataAdapter("select * from " + comboBox1.Text, con);
            try {
                DataTable dt2 = new DataTable();
                da.Fill(dt2);
                dataGridView1.DataSource = dt2;
                con.Close();
            } catch (Exception) {

                return;
            }
 
        }
    }
}
