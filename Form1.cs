using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            DateTime start_date = Convert.ToDateTime("6/8/2018");
            int period = Convert.ToInt32(Math.Ceiling((DateTime.Now - start_date).TotalDays/6));


            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.AllowUserToAddRows = false;

            DataTable dt = new DataTable();
            dt.Columns.Add("SAP Session", typeof(short));
            dt.Columns.Add("Report", typeof(String));
            dt.Columns.Add("Created On Start", typeof(String));
            dt.Columns.Add("Created On End", typeof(String));
            dt.Rows.Add(new object[] { 0, "ZSD_CONT_LIST", start_date.ToString("MM/dd/yyyy"), start_date.AddDays(period).ToString("MM/dd/yyyy") });
            dt.Rows.Add(new object[] { 1, "ZSD_CONT_LIST", start_date.AddDays(period+1).ToString("MM/dd/yyyy"), start_date.AddDays((2*period)+1).ToString("MM/dd/yyyy") });
            dt.Rows.Add(new object[] { 2, "ZSD_CONT_LIST", start_date.AddDays((2 * period) + 2).ToString("MM/dd/yyyy"), start_date.AddDays((3 * period) + 2).ToString("MM/dd/yyyy") });
            dt.Rows.Add(new object[] { 3, "ZSD_CONT_LIST", start_date.AddDays((3 * period) + 3).ToString("MM/dd/yyyy"), start_date.AddDays((4 * period) + 3).ToString("MM/dd/yyyy") });
            dt.Rows.Add(new object[] { 4, "ZSD_CONT_LIST", start_date.AddDays((4 * period) + 4).ToString("MM/dd/yyyy"), start_date.AddDays((5 * period) + 4).ToString("MM/dd/yyyy") });
            dt.Rows.Add(new object[] { 5, "ZSD_CONT_LIST", start_date.AddDays((5 * period) + 5).ToString("MM/dd/yyyy"), start_date.AddDays((6 * period) + 5).ToString("MM/dd/yyyy") });

            DataGridViewComboBoxColumn report = new DataGridViewComboBoxColumn();
            var list11 = new List<string>() { "ZSD_CONT_LIST", "ZSD_REPOMW" };
            report.DataSource = list11;
            report.HeaderText = "Report";
            report.DataPropertyName = "Report";

            DataGridViewTextBoxColumn session = new DataGridViewTextBoxColumn();
            session.HeaderText = "SAP Session";
            session.DataPropertyName = "SAP Session";

            DataGridViewTextBoxColumn start = new DataGridViewTextBoxColumn();
            start.HeaderText = "Created On Start";
            start.DataPropertyName = "Created On Start";
            start.ReadOnly = false;

            DataGridViewTextBoxColumn end = new DataGridViewTextBoxColumn();
            end.HeaderText = "Created On End";
            end.DataPropertyName = "Created On End";
            end.ReadOnly = false;

            dataGridView1.DataSource = dt;
            dataGridView1.Columns.AddRange(session, report,start,end);

            



        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string session_name = row.Cells[0].Value.ToString();
                string report_name = row.Cells[1].Value.ToString();
                string created_on_start_date = row.Cells[2].Value.ToString();
                string created_on_end_date = row.Cells[3].Value.ToString();


                Process p = new Process();


                p.StartInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + report_name + ".exe";
                p.StartInfo.Arguments = created_on_start_date + " " + created_on_end_date + " True " + session_name;
                p.Start();
            }
        }
    }
}
