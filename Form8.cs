using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace WindowsFormsApp1
{
    public partial class Form8 : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        SqlDataAdapter da;
        SqlCommandBuilder builder;
        DataTable dt;
        DataSet ds;
        string id;
        int id1;
        int delete_id;

         public Form8()
        {

            InitializeComponent();
            conn = new SqlConnection();
            conn.ConnectionString = "Integrated Security=True;Persist Security Info=False;Initial Catalog=Exam001;Data Source=DESKTOP-0B1IGGO";
        }
        

        private void Form8_Load(object sender, EventArgs e)
        {   
            
            //da = new SqlDataAdapter("SELECT * FROM Employees", conn);
            //builder = new SqlCommandBuilder(da);
            //dt = new DataTable();
            //da.Fill(dt);
            //da.FillSchema(dt, SchemaType.Source);
            dataGridView1.DataSource = table();

        }

        private void Form8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode.ToString() =="S")
            {
                Savebtn.PerformClick();
            }
            if (e.Control && e.KeyCode.ToString() == "R")
            {
                toolStripButton1.PerformClick();
            }
        }


        public void displayData()
        {
            conn.Open();
            da = new SqlDataAdapter("SELECT * FROM Employees", conn);
            ds = new DataSet();
            builder = new SqlCommandBuilder(da);
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            conn.Close();
     }
        DataTable table() {

            conn.Open();
            da = new SqlDataAdapter("SELECT * FROM Employees", conn);
            dt = new DataTable();
            da.Fill(dt);
            conn.Close();
            return dt;
        }
        private void Savebtn_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            dataGridView1.DataSource = dt;
            builder = new SqlCommandBuilder(da);
            da.Update(dt);
            da = new SqlDataAdapter("SELECT * FROM Employees", conn);
            dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;



            ////////conn.Open();
            ////////cmd = new SqlCommand("UPDATE Employees SET EmployeeID='" + dataGridView1.CurrentRow
            ////////.Cells["employeeIDDataGridViewTextBoxColumn"].Value.ToString() + "'
            ////////, EmployeeName ='" + dataGridView1.CurrentRow.Cells["employeeNameDataGridViewTextBoxColumn"]
            ////////.Value.ToString() + "',BirthDate='" + Convert.ToDateTime(dataGridView1.CurrentRow.Cells
            ////////["birthDateDataGridViewTextBoxColumn"].Value)  + "',Grade='" + dataGridView1.CurrentRow.Cells
            ////////["gradeDataGridViewTextBoxColumn"].Value.ToString() + "',Department='" + dataGridView1.CurrentRow
            ////////.Cells["departmentDataGridViewTextBoxColumn"].Value.ToString() + "',HireDate='" + Convert.ToDateTime
            ////////(dataGridView1.CurrentRow.Cells["hireDateDataGridViewTextBoxColumn"].Value) + "',ProbationPeriod ='"
            ////////+ dataGridView1.CurrentRow.Cells["probationPeriodDataGridViewTextBoxColumn"].Value
            ////////.ToString() + "',Inactive='" + dataGridView1.CurrentRow.Cells["inactiveDataGridViewTextBoxColumn"]
            ////////.Value.ToString() + "',InactivationDate='" +dataGridView1.CurrentRow.Cells["inactivationDateDataGridViewTextBoxColumn"]
            ////////.Value.ToString() + "' WHERE EmployeeID='" + dataGridView1.CurrentRow.Cells["employeeIDDataGridViewTextBoxColumn"]
            ////////.Value.ToString() + "'", conn);
            ////////cmd.ExecuteNonQuery();
            ////////MessageBox.Show("Your Data has been Updated ");
            ////////conn.Close();
            ////////displayData();
            //da.Update(ds.Tables[0]);
            //MessageBox.Show("data updated....");
            //displayData();
            ////int rowIndex = dataGridView1.CurrentCell.RowIndex;
            ////dataGridView1.Rows.Add();
            ////displayData();
            //string ID, nameText, gradeText, dep, probation, inactive, inactivationdate;
            //DateTime birthdate, hiredate;
            //ID = dataGridView1.CurrentRow.Cells["employeeIDDataGridViewTextBoxColumn"].Value.ToString();
            //nameText = dataGridView1.CurrentRow.Cells["employeeNameDataGridViewTextBoxColumn"].Value.ToString();
            //birthdate = Convert.ToDateTime(dataGridView1.CurrentRow.Cells["birthDateDataGridViewTextBoxColumn"].Value);
            //gradeText = dataGridView1.CurrentRow.Cells["gradeDataGridViewTextBoxColumn"].Value.ToString();
            //dep = dataGridView1.CurrentRow.Cells["departmentDataGridViewTextBoxColumn"].Value.ToString();\
            //hiredate = Convert.ToDateTime(dataGridView1.CurrentRow.Cells["hireDateDataGridViewTextBoxColumn"].Value);
            //probation = dataGridView1.CurrentRow.Cells["probationPeriodDataGridViewTextBoxColumn"].Value.ToString();
            //inactive = dataGridView1.CurrentRow.Cells["inactiveDataGridViewTextBoxColumn"].Value.ToString();
            //inactivationdate = dataGridView1.CurrentRow.Cells["inactivationDateDataGridViewTextBoxColumn"].Value.ToString();
            //conn.Open();
            //cmd = new SqlCommand("UPDATE Employees SET EmployeeID='" + ID + "', EmployeeName ='" + nameText + "',BirthDate='" + Convert.ToDateTime(birthdate) + "',Grade='" + gradeText + "',Department='" + dep + "',HireDate='" + Convert.ToDateTime(hiredate) + "',ProbationPeriod ='" + probation + "',Inactive='" + inactive + "',InactivationDate='" + Convert.ToDateTime(inactivationdate) + "' WHERE EmployeeID='" + ID + "'", conn);
            //cmd.ExecuteNonQuery();
            //conn.Close();
            //dataGridView1.DataSource = table();
            //MessageBox.Show("data updated....");



        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Savebtn.PerformClick();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton1.PerformClick();

        }


        private void saveasreshbtn_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Employees info";


            for (int i = 1; i < dataGridView1.Columns.Count+1; i++)
            {
                worksheet.Cells[1, i]= dataGridView1.Columns[i - 1].HeaderText;


            }
            for (int i = 0; i < dataGridView1.Rows.Count ; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1]=dataGridView1.Rows[i].Cells[j].Value;
                }

            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "output";
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.Filter = "All Files| *.*";
            if (saveFileDialog.ShowDialog()==DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing);
            }
            app.Quit();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            conn.Open();
            da = new SqlDataAdapter("SELECT * FROM Employees", conn);
            dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            conn.Close();

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (toolStripProgressBar1.Value <= toolStripProgressBar1.Maximum - 1)

                toolStripProgressBar1.Value += 1;
            else
                timer1.Enabled = false;
               
        }

        private void dataGridView1_CellValueChanged_1(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Savebtn_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text="Save changes ..... ";
        }

        private void Savebtn_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void toolStripButton1_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Refresh data ..... ";
        }

        private void toolStripButton1_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void Refreshbtn_MouseHover(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "save data as exel file  ..... ";
        }

        private void Refreshbtn_MouseLeave(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";

        }
    }
}
