using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataManipulation
{
    public partial class DataManipulation : Form
    {
        private OleDbConnection thisConnection;

        public DataManipulation()
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

        private void DisplayEmployeesWithLeastSalary()
        {
            string queryString = "SELECT EMPLOYEEIDNO as [ID Number], EMPLOYEEFIRSTNAME + ' ' + EMPLOYEELASTNAME as [Full Name], EMPLOYEESALARY as Salary, EMPLOYEEDEPARTMENT as Department FROM EMPLOYEEFILE WHERE EMPLOYEESALARY = (SELECT MIN(EMPLOYEESALARY) FROM EMPLOYEEFILE)";
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, thisConnection);
            DataTable table = new DataTable();
            adapter.Fill(table);

            dataGridView6.DataSource = table;
        }

        private void DisplayEmployeesAndTrainingCourses()
        {
            // Define the SQL query to fetch data from the EMPLOYEEFILE, ATTENDANCEFILE, and TRAININGFILE tables
            string queryString = "SELECT EMPLOYEEFILE.EMPLOYEEIDNO, EMPLOYEEFILE.EMPLOYEEFIRSTNAME, EMPLOYEEFILE.EMPLOYEELASTNAME, TRAININGFILE.TRAININGCOURSE, EMPLOYEEFILE.EMPLOYEEDEPARTMENT FROM ((EMPLOYEEFILE INNER JOIN ATTENDANCEFILE ON EMPLOYEEFILE.EMPLOYEEIDNO = ATTENDANCEFILE.ATTENDANCETRAININGEMPLOYEEID) INNER JOIN TRAININGFILE ON ATTENDANCEFILE.ATTENDANCETRAININGCODE = TRAININGFILE.TRAININGCODE)";

            // Create an OleDbDataAdapter object to fetch data from the database
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, thisConnection);

            // Create a DataTable object to hold the fetched data
            DataTable table = new DataTable();

            // Fill the DataTable with the fetched data using the OleDbDataAdapter
            adapter.Fill(table);

            // Set the data source of the DataGridView control to the filled DataTable
            dataGridView6.DataSource = table;

            // Set the column names in the DataGridView control to match the columns provided
            dataGridView6.Columns[0].HeaderText = "EMPLOYEEIDNO";
            dataGridView6.Columns[1].HeaderText = "EMPLOYEEFIRSTNAME";
            dataGridView6.Columns[2].HeaderText = "TRAININGCOURSE";
            dataGridView6.Columns[3].HeaderText = "EMPLOYEEDEPARTMENT";
        }

        private void DisplayEmployeesWithMultipleTrainings()
        {
            // Define the SQL query to fetch data from the ATTENDANCEFILE and EMPLOYEEFILE tables
            string queryString = "SELECT EMPLOYEELASTNAME as [Last Name], COUNT(*) as [Training Count] FROM ATTENDANCEFILE INNER JOIN EMPLOYEEFILE ON ATTENDANCEFILE.ATTENDANCETRAININGEMPLOYEEID = EMPLOYEEFILE.EMPLOYEEIDNO GROUP BY EMPLOYEELASTNAME HAVING COUNT(*) > 1";

            // Create an OleDbDataAdapter object to fetch data from the database
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, thisConnection);

            // Create a DataTable object to hold the fetched data
            DataTable table = new DataTable();

            // Fill the DataTable with the fetched data using the OleDbDataAdapter
            adapter.Fill(table);

            // Set the data source of the DataGridView control to the filled DataTable
            dataGridView6.DataSource = table;

            // Set the column names in the DataGridView control
            dataGridView6.Columns[0].HeaderText = "Last Name";
            dataGridView6.Columns[1].HeaderText = "Training Count";
        }

        private void DisplayParticipantsOfAngerManagementTraining()
        {
            // Define the SQL query to fetch data from the EMPLOYEEFILE, ATTENDANCEFILE, and TRAININGFILE tables
            string queryString = "SELECT EMPLOYEEFILE.EMPLOYEEFIRSTNAME + ' ' + EMPLOYEEFILE.EMPLOYEELASTNAME AS [Full Name], EMPLOYEEFILE.EMPLOYEEDEPARTMENT AS [Department] " +
                                 "FROM ((ATTENDANCEFILE " +
                                 "INNER JOIN EMPLOYEEFILE ON ATTENDANCEFILE.ATTENDANCETRAININGEMPLOYEEID = EMPLOYEEFILE.EMPLOYEEIDNO) " +
                                 "INNER JOIN TRAININGFILE ON ATTENDANCEFILE.ATTENDANCETRAININGCODE = TRAININGFILE.TRAININGCODE) " +
                                 "WHERE TRAININGFILE.TRAININGCOURSE = 'ANGER MANAGEMENT'";

            // Create an OleDbDataAdapter object to fetch data from the database
            OleDbDataAdapter adapter = new OleDbDataAdapter(queryString, thisConnection);

            // Create a DataTable object to hold the fetched data
            DataTable table = new DataTable();

            // Fill the DataTable with the fetched data using the OleDbDataAdapter
            adapter.Fill(table);

            // Set the data source of the DataGridView control to the filled DataTable
            dataGridView6.DataSource = table;

            // Set the column names in the DataGridView control
            dataGridView6.Columns[0].HeaderText = "Full Name";
            dataGridView6.Columns[1].HeaderText = "Department";
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

        private void button6_Click(object sender, EventArgs e)
        {
            DisplayEmployeesWithLeastSalary();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DisplayEmployeesAndTrainingCourses();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DisplayEmployeesWithMultipleTrainings();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DisplayParticipantsOfAngerManagementTraining();
        }
    }
}