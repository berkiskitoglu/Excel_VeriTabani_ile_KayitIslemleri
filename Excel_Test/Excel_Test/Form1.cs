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

namespace Excel_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //C:\Users\iskit\Desktop\Bil-2012-2013-Yaz-Ders-Prog-2.xlsx
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\iskit\Desktop\Bil-2012-2013-Yaz-Ders-Prog-2.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;'");

        void listele() 
        {
            baglanti.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("Select * from [Sayfa1$]", baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            baglanti.Close();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            listele();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("insert into [Sayfa1$] (Saat,Ders) values(@p1,@p2)", baglanti);
            komut.Parameters.AddWithValue("@p1", textBox1.Text);
            komut.Parameters.AddWithValue("@p2", textBox2.Text);
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Bağlantı Başarılı Bir Şekilde Gerçekleşti");
        }
    }
}
