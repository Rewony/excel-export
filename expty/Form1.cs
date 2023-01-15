using System.Data;
using System.Data.SqlClient;
using Microsoft.Data.SqlClient;



namespace expty
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            SqlConnection c = new SqlConnection("Data Source=.;Initial Catalog=ProgrammindDB;Integrated Security=True;Encrypt=False;");

            SqlCommand CONNECTÝON = new SqlCommand("SELECT * FROM product_setup_table", c);
            SqlDataAdapter d = new SqlDataAdapter(CONNECTÝON);
            DataTable dt = new DataTable();
            d.Fill(dt);
            dataGridView1.DataSource = dt;






        }

        private void button2_Click(object sender, EventArgs e)
        {

            dataGridView1.SelectAll();

            DataObject copydata = dataGridView1.GetClipboardContent();
            if(copydata != null)
            {

                Clipboard.SetDataObject(copydata);

                Microsoft.Office.Interop.Excel.Application xlapp= new Microsoft.Office.Interop.Excel.Application();

                xlapp.Visible = true;

                Microsoft.Office.Interop.Excel.Workbook xlWbook;
                Microsoft.Office.Interop.Excel.Worksheet xlsheet;
                object miseddata = System.Reflection.Missing.Value;
                xlWbook = xlapp.Workbooks.Add(miseddata);

                xlsheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWbook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range xlr = (Microsoft.Office.Interop.Excel.Range)xlsheet.Cells[1, 1];
                xlr.Select();
                xlsheet.PasteSpecial(xlr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);




            }




        }
    }
}