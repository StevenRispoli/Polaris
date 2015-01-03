using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Data;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;

namespace Polaris
{
    /// <summary>
    /// Interaction logic for Login.xaml
    /// </summary>
    
    
    public partial class Login : Window
    {
        internal event EventHandler loginSuccessful;
        Registration registration = new Registration();
        string adminEmail;
        string adminFirstName;
        string adminLastName;
        
        static string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Polaris.accdb;Persist Security Info=False;";
        OleDbConnection dbConnection = new OleDbConnection(connectionString);
        
        public Login()
        {
            InitializeComponent();
            getAdminInfo();
        }

        private void getAdminInfo()
        {
            dbConnection.Open();
            string adminInfoQuery = "SELECT * FROM tblUsers WHERE security_Level = 4";
            OleDbCommand checkAdminInfo = new OleDbCommand(adminInfoQuery, dbConnection);
            checkAdminInfo.ExecuteNonQuery();
            OleDbDataReader dataReader = checkAdminInfo.ExecuteReader();
            try
            {
                int count = 0;
                while (dataReader.Read())
                {
                    count++;
                    if (count == 1)
                    {
                        adminEmail = Convert.ToString(dataReader["email"]);
                        adminFirstName = Convert.ToString(dataReader["first_Name"]);
                        adminLastName = Convert.ToString(dataReader["last_Name"]);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbConnection.Close();
        }
        
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            string loginQuery = "SELECT * FROM tblUsers WHERE username = '" + this.txtUsername.Text + "'";
            OleDbCommand checkUsername = new OleDbCommand(loginQuery, dbConnection);
            checkUsername.ExecuteNonQuery();
            OleDbDataReader dataReader = checkUsername.ExecuteReader();
            try
            {
                int count = 0;
                while (dataReader.Read())
                {
                    count++;
                    
                    if (count == 1)
                    {
                        int loginAttempts = Convert.ToInt32(dataReader["login_Attempts"]);
                        if (loginAttempts < 3)
                        {
                            string enteredPassword = this.txtPassword.Password;
                            string validPassword = Convert.ToString(dataReader["password"]);
                            if (enteredPassword == validPassword)
                            {
                                string resetAttemptsQuery = "UPDATE tblUsers SET login_Attempts = 0 WHERE username = '" + this.txtUsername.Text + "'";
                                OleDbCommand resetAttempts = new OleDbCommand(resetAttemptsQuery, dbConnection);
                                resetAttempts.ExecuteNonQuery();
                                Security.securityLevel = Convert.ToInt32(dataReader["security_Level"]);
                                Security.employeeID = Convert.ToInt32(dataReader["employee_ID"]);
                                loginSuccessful(this, null);
                                Close();
                            }
                            else
                            {
                                string attemptsQuery = "UPDATE tblUsers SET login_Attempts = '" + (loginAttempts += 1) + "' WHERE username = '" + this.txtUsername.Text + "'";
                                OleDbCommand updateAttempts = new OleDbCommand(attemptsQuery, dbConnection);
                                updateAttempts.ExecuteNonQuery();
                                if (loginAttempts < 3)
                                {
                                    MessageBox.Show("Incorrect Password");
                                    break;
                                }
                                else
                                {
                                    //Based off of user "Domenic" at this address: http://stackoverflow.com/questions/32260/sending-email-in-net-through-gmail
                                    MessageBox.Show("You have attempted to log in too many times. Please contact an administrator.");
                                    var fromAddress = new MailAddress("nexusmedicalsystems@gmail.com", "Nexus Medical Systems Inc.");
                                    var toAddress = new MailAddress(adminEmail, "'" + adminFirstName + " " + adminLastName + "'");
                                    const string fromPassword = "nextgenmed";
                                    const string subject = "Registered user failed to log in";
                                    string body = "" + Convert.ToString(dataReader["first_Name"]) + " " + Convert.ToString(dataReader["last_Name"]) + " has failed to log in too many times. Please reset this user's login attempts";

                                    var smtp = new SmtpClient
                                    {
                                        Host = "smtp.gmail.com",
                                        Port = 587,
                                        EnableSsl = true,
                                        DeliveryMethod = SmtpDeliveryMethod.Network,
                                        Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
                                        //Timeout = 900
                                    };
                                    using (var message = new MailMessage(fromAddress, toAddress)
                                    {
                                        Subject = subject,
                                        Body = body
                                    })
                                    {
                                        smtp.Send(message);
                                    }
                                    Close();
                                }    
                            }
                        }
                        else
                        {
                            MessageBox.Show("You have attempted to log in too many times. Please contact an administrator.");
                            Close();
                        }
                    }
                }
                if(count < 1)
                {
                    MessageBox.Show("Invalid Username");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            dbConnection.Close();
        }

        
        
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            registration.Show();
        }
    }
}

