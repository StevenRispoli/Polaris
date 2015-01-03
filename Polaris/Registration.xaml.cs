using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Data.OleDb;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Polaris
{
    /// <summary>
    /// Interaction logic for Registration.xaml
    /// </summary>
    public partial class Registration : Window
    {
        static string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Polaris.accdb;Persist Security Info=False;";
        OleDbConnection dbConnection = new OleDbConnection(connectionString);
        public Registration()
        {
            InitializeComponent();
        }

        private void btnSubmitRegistration_Click(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdUsersInsert = new OleDbCommand("INSERT INTO tblUsers(last_Name, first_Name, username, [email], [password], security_Level, login_Attempts) VALUES("
                + "@last_Name, "
                + "@first_Name, "
                + "@username, "
                + "@email, "
                + "@password, "
                + "@security_Level, "
                + "@login_Attempts)", dbConnection);
                cmdUsersInsert.Parameters.Add(new OleDbParameter("@last_Name", this.txtLast_NameRegister.Text));
                cmdUsersInsert.Parameters.Add(new OleDbParameter("@first_Name", this.txtFirst_NameRegister.Text));
                cmdUsersInsert.Parameters.Add(new OleDbParameter("@username", this.txtUsernameRegister.Text));
                cmdUsersInsert.Parameters.Add(new OleDbParameter("@email", this.txtEmailRegister.Text));
                cmdUsersInsert.Parameters.Add(new OleDbParameter("@password", this.txtPasswordRegister.Password.ToString()));
                cmdUsersInsert.Parameters.Add(new OleDbParameter("@security_Level", Convert.ToInt32(this.txtSecurity_LevelRegister.Text)));
                cmdUsersInsert.Parameters.Add(new OleDbParameter("@login_Attempts", Convert.ToInt32(this.txtLogin_AttemptsRegister.Text)));
                cmdUsersInsert.ExecuteNonQuery();
                MessageBox.Show("Save Successful");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbConnection.Close();
            this.Hide();
        }
    }
}
