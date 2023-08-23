using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace dubletter_patienter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Datacontainer.connectsource = "Data Source=Klingen-su-db,62468; Initial Catalog = Klingen;";
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            Datacontainer.connectsource = "Data Source=Klingen-su-db,62468; Initial Catalog = Klingen;";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            Datacontainer.connectsource = "Data Source=Klingen-test-su-db,62468; Initial Catalog = Klingen_Test;";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Datacontainer.anvandarnamn = textBox1.Text;
            Datacontainer.losen = textBox2.Text;
            Datacontainer.connectionString = @Datacontainer.connectsource + "User ID=" + textBox1.Text + ";Password=" + textBox2.Text + "";
            Datacontainer.cnn = new SqlConnection(Datacontainer.connectionString);
            Datacontainer.cnn.Open();
            //   MessageBox.Show("Connection Open  !");
            string message = "Connection Open  !";
            string title = "";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.OK)
            {
                this.Close();
            }
            else
            {
                // Do something
            }



        }
    }
}
