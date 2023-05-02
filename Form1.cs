using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataManipulation
{
    public partial class Form1 : Form
    {
        private OleDbConnection thisConnection;

        public Form1()
        {
            InitializeComponent();
            thisConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\solis\\source\\repos\\DataManipulation\\KOMPANYA.accdb");
        }

        private void DisplayEmployeesAboveProductionAvgSalary()
        {
            string queryString = "SELECT EMPLOYEEIDNO, EMPLOYEEFIRSTNAME + ' ' + EMPLOYEELASTNAME as FULLNAME, EMPLOYEEDEPARTMENT, EMPLOYEESALARY FROM EMPLOYEEFILE WHERE EMPLOYEESALARY > (SELECT AVG(EMPLOYEESALARY) FROM EMPLOYEEFILE WHERE EMPLOYEEDEPARTMENT = 'PRODUCTION')";
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, thisConnection);
            DataSet dataSet = new DataSet();
            try
            {
                thisConnection.Open();
                adapter.Fill(dataSet);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                thisConnection.Close();
            }

            if (dataSet.Tables.Count > 0)
            {
                dataGridView6.DataSource = dataSet.Tables[0];
            }
            else
            {
                Console.WriteLine("No tables found in DataSet");
            }
        }

        private void DisplayTotalSalaries()
        {
            string queryString = "SELECT SUM(EMPLOYEESALARY) FROM EMPLOYEEFILE";
            OleDbCommand command = new OleDbCommand(queryString, thisConnection);
            thisConnection.Open();
            object result = command.ExecuteScalar();
            thisConnection.Close();

            if (result != DBNull.Value && result != null)
            {
                int totalSalary = Convert.ToInt32(result);
                MessageBox.Show(string.Format("Total salaries of KOMPANYA INC.: {0:C}", totalSalary));
            }
            else
            {
                MessageBox.Show("No data found");
            }
        }

        private void DisplayAvgSalaryOfLastNameStartingWithM()
        {
            string queryString = "SELECT AVG(EMPLOYEESALARY) FROM EMPLOYEEFILE WHERE EMPLOYEELASTNAME LIKE 'M%'";
            OleDbCommand command = new OleDbCommand(queryString, thisConnection);
            thisConnection.Open();
            object result = command.ExecuteScalar();
            thisConnection.Close();

            if (result != DBNull.Value && result != null)
            {
                double averageSalary = Convert.ToDouble(result);
                MessageBox.Show(string.Format("Average salary of employees whose last name starts with 'M': {0:C}", averageSalary));
            }
            else
            {
                MessageBox.Show("No data found");
            }
        }

        private void DisplayProductionEmployeesAboveDeptAvgSalary()
        {
            string queryString = "SELECT EMPLOYEEIDNO as [ID NUMBER], EMPLOYEEFIRSTNAME as FIRSTNAME, EMPLOYEELASTNAME as LASTNAME, EMPLOYEESALARY as SALARY FROM EMPLOYEEFILE WHERE EMPLOYEEDEPARTMENT = 'PRODUCTION' AND EMPLOYEESALARY > (SELECT AVG(EMPLOYEESALARY) FROM EMPLOYEEFILE WHERE EMPLOYEEDEPARTMENT = 'PRODUCTION')";
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, thisConnection);
            DataTable table = new DataTable();
            adapter.Fill(table);


            // Set the data source of the DataGridView control to the new table
            dataGridView6.DataSource = table;
        }

        private void DisplayTotalSalesDepartmentSalaries()
        {
            string queryString = "SELECT SUM(EMPLOYEESALARY) FROM EMPLOYEEFILE WHERE EMPLOYEEDEPARTMENT = 'SALES'";
            OleDbCommand command = new OleDbCommand(queryString, thisConnection);
            thisConnection.Open();
            object result = command.ExecuteScalar();
            thisConnection.Close();

            if (result != DBNull.Value && result != null)
            {
                int totalSalary = Convert.ToInt32(result);
                MessageBox.Show(string.Format("Total salaries of SALES department employees: {0:C}", totalSalary));
            }
            else
            {
                MessageBox.Show("No data found");
            }

            // Alternatively, you can display the result in a label control:
            // label1.Text = string.Format("Total salaries of SALES department employees: {0:C}", totalSalary);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DisplayEmployeesAboveProductionAvgSalary();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DisplayTotalSalaries();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DisplayAvgSalaryOfLastNameStartingWithM();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DisplayProductionEmployeesAboveDeptAvgSalary();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DisplayTotalSalesDepartmentSalaries();
        }
    }
}
