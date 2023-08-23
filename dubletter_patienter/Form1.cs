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
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

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
            string message = "Connection Open  !";
            string title = "";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.OK)
            {
                button2.Enabled = true;
            }
            else
            {
                // Do something
            }



        }

        private void button2_Click(object sender, EventArgs e)
        {

            // Start a new workbook in Excel.
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            oXL = new Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = "Index";
            oSheet.Cells[1, 2] = "Personnummer";
            oSheet.Cells[1, 3] = "Familyname";
            oSheet.Cells[1, 4] = "Förnamn";

            String Sql;
            Sql = "SELECT ROW_NUMBER() OVER(ORDER BY[Index] Desc) AS RowNumber,[Index],[Personal number],[Familyname],[First Name] FROM[Klingen].[dbo].[Patients] WHERE[Personal number] IN(SELECT[Personal number] FROM[Klingen].[dbo].[Patients] GROUP BY[Personal number] HAVING COUNT(*) > 1)";
            Datacontainer.command = new SqlCommand(Sql, Datacontainer.cnn);
            Datacontainer.command.CommandType = CommandType.Text;
            SqlDataReader reader = Datacontainer.command.ExecuteReader();
            int radnummer;
            radnummer = 1;
            while (reader.Read())
            {

                Datacontainer.Index = (int)reader.GetValue(1);
                Datacontainer.personnummer = (String)reader.GetValue(2);
                Datacontainer.Familyname = (String)reader.GetValue(3);
                Datacontainer.fornamn = (String)reader.GetValue(4);

                //För nu över till excel!
                oSheet.Cells[radnummer,1] = Datacontainer.Index;
                oSheet.Cells[radnummer,2] = Datacontainer.personnummer;
                oSheet.Cells[radnummer,3] = Datacontainer.Familyname;
                oSheet.Cells[radnummer,4] = Datacontainer.fornamn;
                radnummer++;

            }

            reader.Close();

        }
    }
}
