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
using System.Windows.Navigation;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Diagnostics.Eventing;
using System.Xaml;
using System.Xaml.Permissions;
using System.Collections;
using System.Collections.ObjectModel;
using System.Printing;

namespace Polaris
{
    //Saves value of logged in user's security level
    public static class Security
    {
        public static int securityLevel;
        public static int employeeID;
    }
    
    
    public partial class MainWindow : Window
    {
        static string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Polaris.accdb;Persist Security Info=False;";

        OleDbConnection dbConnection = new OleDbConnection(connectionString);
        DataSet dataSet;
        
        public MainWindow()
        {
            InitializeComponent();
            refreshComboBoxes();
        }

        private void refreshComboBoxes()
        {
            fillComboBox("SELECT * FROM tblComboLocation", cmbPlate_Location);
            fillComboBox("SELECT * FROM tblComboSize", cmbPlate_Size);
            fillComboBox("SELECT * FROM tblComboLocation", cmbCassette_Location);
            fillComboBox("SELECT * FROM tblComboSize", cmbCassette_Size);
            fillComboBox("SELECT * FROM tblComboSize", cmbDeleteFromComboSizes);
            fillComboBox("SELECT * FROM tblComboLocation", cmbDeleteFromComboLocations);
        }

        //Shows MainWindow after successful login
        internal void startupMainWindow(object sender, EventArgs e)
        {
            Application.Current.MainWindow = this;
            Show(); 
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //Set security level
            if (Security.securityLevel < 4)
            {
                tabUsers.Visibility = Visibility.Collapsed;
                tabUsers.Visibility = Visibility.Collapsed;
            }
            if (Security.securityLevel < 3)
            {
                tabUsers.Visibility = Visibility.Collapsed;
                tabUsers.Visibility = Visibility.Collapsed;
                btnDeletePlate.Visibility = Visibility.Collapsed;
                btnDeleteCassette.Visibility = Visibility.Collapsed;
                btnDeleteCassetteCleaning.Visibility = Visibility.Collapsed;
                btnDeleteCassetteServicing.Visibility = Visibility.Collapsed;
                btnDeletePlateCleaning.Visibility = Visibility.Collapsed;
                btnDeleteUser.Visibility = Visibility.Collapsed;

            }
            if (Security.securityLevel < 2)
            {
                tabUsers.Visibility = Visibility.Collapsed;
                tabUsers.Visibility = Visibility.Collapsed;
                btnDeletePlate.Visibility = Visibility.Collapsed;
                btnDeleteCassette.Visibility = Visibility.Collapsed;
                btnDeleteCassetteCleaning.Visibility = Visibility.Collapsed;
                btnDeleteCassetteServicing.Visibility = Visibility.Collapsed;
                btnDeletePlateCleaning.Visibility = Visibility.Collapsed;
                btnDeleteUser.Visibility = Visibility.Collapsed;
                btnUpdateCassette.Visibility = Visibility.Collapsed;
                btnUpdateCassetteCleaning.Visibility = Visibility.Collapsed;
                btnUpdateCassetteServicing.Visibility = Visibility.Collapsed;
                btnUpdatePlate.Visibility = Visibility.Collapsed;
                btnUpdatePlateCleaning.Visibility = Visibility.Collapsed;
                btnUpdateUser.Visibility = Visibility.Collapsed;
            }
            if (Security.securityLevel < 1)
            {
                tabUsers.Visibility = Visibility.Collapsed;
                tabUsers.Visibility = Visibility.Collapsed;
                btnDeletePlate.Visibility = Visibility.Collapsed;
                btnDeleteCassette.Visibility = Visibility.Collapsed;
                btnDeleteCassetteCleaning.Visibility = Visibility.Collapsed;
                btnDeleteCassetteServicing.Visibility = Visibility.Collapsed;
                btnDeletePlateCleaning.Visibility = Visibility.Collapsed;
                btnDeleteUser.Visibility = Visibility.Collapsed;
                btnUpdateCassette.Visibility = Visibility.Collapsed;
                btnUpdateCassetteCleaning.Visibility = Visibility.Collapsed;
                btnUpdateCassetteServicing.Visibility = Visibility.Collapsed;
                btnUpdatePlate.Visibility = Visibility.Collapsed;
                btnUpdatePlateCleaning.Visibility = Visibility.Collapsed;
                btnUpdateUser.Visibility = Visibility.Collapsed;
                //filePrint.Visibility = Visibility.Collapsed;
            }
            
            try
            {
                dbConnection.Open();
                OleDbCommand checkPlateCleanings = new OleDbCommand("SELECT * FROM tblPlates WHERE in_Use = YES", dbConnection);
                checkPlateCleanings.ExecuteNonQuery();
                OleDbDataReader dataReader = checkPlateCleanings.ExecuteReader();
                string message = "";
                while(dataReader.Read())
                {
                    string plate_Barcode = Convert.ToString(dataReader["plate_Barcode"]);
                    message = ""+ plate_Barcode +"\r\n"+ message +"";
                }
                MessageBox.Show("The following plates need to be cleaned:\r\n\n" + message);
                dbConnection.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                dbConnection.Close();
            }
            
            try
            {
                dbConnection.Open();
                OleDbCommand checkCassetteCleanings = new OleDbCommand("SELECT * FROM tblCassettes WHERE in_Use = YES", dbConnection);
                checkCassetteCleanings.ExecuteNonQuery();
                OleDbDataReader dataReader = checkCassetteCleanings.ExecuteReader();
                string message = "";
                while (dataReader.Read())
                {
                    string cassette_Barcode = Convert.ToString(dataReader["cassette_Barcode"]);
                    message = "" + cassette_Barcode + "\r\n" + message + "";
                }
                MessageBox.Show("The following cassettes need to be cleaned:\r\n\n" + message);
                dbConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                dbConnection.Close();
            }
            
            try
            {
                dbConnection.Open();
                OleDbCommand checkCassetteServicings = new OleDbCommand("SELECT * FROM tblCassettes WHERE (((tblCassettes.[needs_Service])=Yes) AND ((tblCassettes.[sent_For_Service]) Is Null))", dbConnection);
                checkCassetteServicings.ExecuteNonQuery();
                OleDbDataReader dataReader = checkCassetteServicings.ExecuteReader();
                string message = "";
                string cassette_Barcode = "";
                while (dataReader.Read())
                {
                    cassette_Barcode = Convert.ToString(dataReader["cassette_Barcode"]);
                    message = "" + cassette_Barcode + "\r\n" + message + "";
                }

                if (string.IsNullOrEmpty(cassette_Barcode))
                    message = " ";
                else
                    MessageBox.Show("The following cassettes need to be sent out for servicing:\r\n\n" + message);
                dbConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                dbConnection.Close();
            }
            dbConnection.Close();
            Polaris.PolarisDataSet PolarisDataSet = ((Polaris.PolarisDataSet)(this.FindResource("PolarisDataSet")));
            // Load data into tblUsers.
            Polaris.PolarisDataSetTableAdapters.tblUsersTableAdapter PolarisDataSettblUsersTableAdapter = new Polaris.PolarisDataSetTableAdapters.tblUsersTableAdapter();
            PolarisDataSettblUsersTableAdapter.Fill(PolarisDataSet.tblUsers);
            System.Windows.Data.CollectionViewSource tblUsersViewSource = ((System.Windows.Data.CollectionViewSource)(this.FindResource("tblUsersViewSource")));
            tblUsersViewSource.View.MoveCurrentToFirst();
            // Load data into tblPlates.
            Polaris.PolarisDataSetTableAdapters.tblPlatesTableAdapter PolarisDataSettblPlatesTableAdapter = new Polaris.PolarisDataSetTableAdapters.tblPlatesTableAdapter();
            PolarisDataSettblPlatesTableAdapter.Fill(PolarisDataSet.tblPlates);
            System.Windows.Data.CollectionViewSource tblPlatesViewSource = ((System.Windows.Data.CollectionViewSource)(this.FindResource("tblPlatesViewSource")));
            tblPlatesViewSource.View.MoveCurrentToFirst();
            // Load data into tblComboSize
            Polaris.PolarisDataSetTableAdapters.tblComboSizeTableAdapter PolarisDataSettblComboSizeTableAdapter = new Polaris.PolarisDataSetTableAdapters.tblComboSizeTableAdapter();
            PolarisDataSettblComboSizeTableAdapter.Fill(PolarisDataSet.tblComboSize);
            System.Windows.Data.CollectionViewSource tblComboSizeViewSource = ((System.Windows.Data.CollectionViewSource)(this.FindResource("tblComboSizeViewSource")));
            tblComboSizeViewSource.View.MoveCurrentToFirst();
            // Load data into tblComboLocation
            Polaris.PolarisDataSetTableAdapters.tblComboLocationTableAdapter PolarisDataSettblComboLocationTableAdapter = new Polaris.PolarisDataSetTableAdapters.tblComboLocationTableAdapter();
            PolarisDataSettblComboLocationTableAdapter.Fill(PolarisDataSet.tblComboLocation);
            System.Windows.Data.CollectionViewSource tblComboLocationViewSource = ((System.Windows.Data.CollectionViewSource)(this.FindResource("tblComboLocationViewSource")));
            tblComboLocationViewSource.View.MoveCurrentToFirst();
            // Load data into tblPlateCleanings.
            Polaris.PolarisDataSetTableAdapters.tblPlateCleaningsTableAdapter PolarisDataSettblPlateCleaningsTableAdapter = new Polaris.PolarisDataSetTableAdapters.tblPlateCleaningsTableAdapter();
            PolarisDataSettblPlateCleaningsTableAdapter.Fill(PolarisDataSet.tblPlateCleanings);
            System.Windows.Data.CollectionViewSource tblPlateCleaningsViewSource = ((System.Windows.Data.CollectionViewSource)(this.FindResource("tblPlateCleaningsViewSource")));
            tblPlateCleaningsViewSource.View.MoveCurrentToFirst();
            // Load data into tblCassettes.
            Polaris.PolarisDataSetTableAdapters.tblCassettesTableAdapter PolarisDataSettblCassettesTableAdapter = new Polaris.PolarisDataSetTableAdapters.tblCassettesTableAdapter();
            PolarisDataSettblCassettesTableAdapter.Fill(PolarisDataSet.tblCassettes);
            System.Windows.Data.CollectionViewSource tblCassettesViewSource = ((System.Windows.Data.CollectionViewSource)(this.FindResource("tblCassettesViewSource")));
            tblCassettesViewSource.View.MoveCurrentToFirst();
            // Load data into tblCassetteCleanings.
            Polaris.PolarisDataSetTableAdapters.tblCassetteCleaningsTableAdapter PolarisDataSettblCassetteCleaningsTableAdapter = new Polaris.PolarisDataSetTableAdapters.tblCassetteCleaningsTableAdapter();
            PolarisDataSettblCassetteCleaningsTableAdapter.Fill(PolarisDataSet.tblCassetteCleanings);
            System.Windows.Data.CollectionViewSource tblCassetteCleaningsViewSource = ((System.Windows.Data.CollectionViewSource)(this.FindResource("tblCassetteCleaningsViewSource")));
            tblCassetteCleaningsViewSource.View.MoveCurrentToFirst();
            // Load data into tblCassetteServicings.
            Polaris.PolarisDataSetTableAdapters.tblCassetteServicingsTableAdapter PolarisDataSettblCassetteServicingsTableAdapter = new Polaris.PolarisDataSetTableAdapters.tblCassetteServicingsTableAdapter();
            PolarisDataSettblCassetteServicingsTableAdapter.Fill(PolarisDataSet.tblCassetteServicings);
            System.Windows.Data.CollectionViewSource tblCassetteServicingsViewSource = ((System.Windows.Data.CollectionViewSource)(this.FindResource("tblCassetteServicingsViewSource")));
            tblCassetteServicingsViewSource.View.MoveCurrentToFirst();
        }
        
        //Fills comboboxs for plates
        private void fillComboBox(string fillQuery, ComboBox comboBox)
        {
            try
            {
                dbConnection.Open();
                OleDbCommand checkLocation = new OleDbCommand(fillQuery, dbConnection);
                OleDbDataReader dataReader = checkLocation.ExecuteReader();
                while (dataReader.Read())
                {
                    string location = dataReader.GetString(0);
                    comboBox.Items.Add(location);
                }
                dbConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dbConnection.Close();
            }
        }

        private void clearComboBox()
        {
            cmbPlate_Location.Items.Clear();
            cmbPlate_Size.Items.Clear();
            cmbCassette_Location.Items.Clear();
            cmbCassette_Size.Items.Clear();
            cmbDeleteFromComboSizes.Items.Clear();
            cmbDeleteFromComboLocations.Items.Clear();
            cmbInputCombobox.Items.Clear();
        }
        
        //Turns text into hyperlink and opens local file that the path points to
        private void DG_Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            Hyperlink link = e.Source as Hyperlink;
            Process.Start(link.NavigateUri.LocalPath);
        }

        //Refreshes datagrid after a sql command is executed
        OleDbDataAdapter refresh = new OleDbDataAdapter();
        public DataSet getDataSet(string sqlCommand, string tblName)
        {
            refresh = new OleDbDataAdapter(sqlCommand, dbConnection);
            dataSet = new DataSet();
            refresh.Fill(dataSet, tblName);
            return dataSet;
        }
        
        private void Button_Click_openAddSize(object sender, RoutedEventArgs e)
        {
            addToComboSizes.Visibility = Visibility.Visible;
        }

        private void Button_Click_closeAddSize(object sender, RoutedEventArgs e)
        {
            addToComboSizes.Visibility = Visibility.Collapsed;
        }
        
        private void Button_Click_openDeleteSize(object sender, RoutedEventArgs e)
        {
            deleteFromComboSizes.Visibility = Visibility.Visible;
        }

        private void Button_Click_closeDeleteSize(object sender, RoutedEventArgs e)
        {
            deleteFromComboSizes.Visibility = Visibility.Collapsed;
        }

        private void Button_Click_openAddLocation(object sender, RoutedEventArgs e)
        {
            addToComboLocations.Visibility = Visibility.Visible;
        }

        private void Button_Click_closeAddLocation(object sender, RoutedEventArgs e)
        {
            addToComboLocations.Visibility = Visibility.Collapsed;
        }

        private void Button_Click_openDeleteLocation(object sender, RoutedEventArgs e)
        {
            deleteFromComboLocations.Visibility = Visibility.Visible;
        }

        private void Button_Click_closeDeleteLocation(object sender, RoutedEventArgs e)
        {
            deleteFromComboLocations.Visibility = Visibility.Collapsed;
        }
        
        private void Button_Click_addSize(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdAddSize = new OleDbCommand("INSERT INTO tblComboSize VALUES(@size)", dbConnection);
                cmdAddSize.Parameters.Add(new OleDbParameter("@size", this.txtAddToComboSizes.Text));
                if (MessageBox.Show("Are you sure you want to add this size?", "Update Plate/Cassette Sizes", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdAddSize.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dbConnection.Close();
                clearComboBox();
                refreshComboBoxes();
                dbConnection.Close();
            }
        }

        private void Button_Click_addLocation(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdAddLocation = new OleDbCommand("INSERT INTO tblComboLocation VALUES (@location)", dbConnection);
                cmdAddLocation.Parameters.Add(new OleDbParameter("@location", this.txtAddToComboLocations.Text));
                if (MessageBox.Show("Are you sure you want to add this location?", "Update Locations", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdAddLocation.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dbConnection.Close();
                clearComboBox();
                refreshComboBoxes();
                dbConnection.Close();
            }
        }

        private void Button_Click_deleteSize(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdDeleteSize = new OleDbCommand("DELETE * FROM tblComboSize WHERE size = ?", dbConnection);
                cmdDeleteSize.Parameters.Add(new OleDbParameter("@size", this.cmbDeleteFromComboSizes.Text));
                if (MessageBox.Show("Are you sure you want to delete this size?", "Delete Plate/Cassette Size", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {

                    cmdDeleteSize.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dbConnection.Close();
                clearComboBox();
                refreshComboBoxes();
                dbConnection.Close();
            }
        }

        private void Button_Click_deleteLocation(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdDeleteLocation = new OleDbCommand("DELETE * FROM tblComboLocation WHERE location = ?", dbConnection);
                cmdDeleteLocation.Parameters.Add(new OleDbParameter("@location", this.cmbDeleteFromComboLocations.Text));
                if (MessageBox.Show("Are you sure you want to delete this location?", "Delete Location", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdDeleteLocation.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dbConnection.Close();
                clearComboBox();
                refreshComboBoxes();
                dbConnection.Close();
            }
        }
        
        private void Button_Click_openUpdateAccountBoxGrid(object sender, RoutedEventArgs e)
        {
            updateAccountBoxGrid.Visibility = Visibility.Visible;
        }

        private void Button_Click_closerUpdateAccountBoxGrid(object sender, RoutedEventArgs e)
        {
            updateAccountBoxGrid.Visibility = Visibility.Collapsed;
        }

        private void Button_Click_updateAccount(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdPasswordUpdate = new OleDbCommand("UPDATE tblUsers SET "
                + "last_Name = @last_Name, "
                + "first_Name = @first_Name, "
                + "[email] = @email, "
                + "[password] = @password "
                + "WHERE employee_ID = @employee_ID", dbConnection);
                cmdPasswordUpdate.Parameters.Add(new OleDbParameter("@last_Name", this.txtLastNameUpdate.Text));
                cmdPasswordUpdate.Parameters.Add(new OleDbParameter("@first_Name", this.txtFirstNameUpdate.Text));
                cmdPasswordUpdate.Parameters.Add(new OleDbParameter("@email", this.txtEmailUpdate.Text));
                cmdPasswordUpdate.Parameters.Add(new OleDbParameter("@password", this.txtPasswordUpdate.Password.ToString()));
                cmdPasswordUpdate.Parameters.Add(new OleDbParameter("@employee_ID", Convert.ToInt32(Security.employeeID)));
                
                if (MessageBox.Show("Are you sure you want to update your account?", "Update User Account", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdPasswordUpdate.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblUsers", "tblUsers");
                tblUsersDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        //Inserts new plate into tblPlates
        private void Button_Click_btnInsertPlate(object sender, RoutedEventArgs e)
        {
            
            try
            {
                dbConnection.Open();
                OleDbCommand cmdPlateInsert = new OleDbCommand("INSERT INTO tblPlates VALUES("
                + "@plate_Barcode, "
                + "@plate_Size, "
                + "@in_Use, "
                + "@in_Use_By, "
                + "@uses_Today, "
                + "@total_Uses, "
                + "@overdue_Cleaning, "
                + "@location, "
                + "@plate_Photo, "
                + "@purchase_Date, "
                + "@warranty_Expire, "
                + "@out_Of_Service, "
                + "@date_Replaced, "
                + "@notes)", dbConnection);
                cmdPlateInsert.Parameters.Add(new OleDbParameter("@plate_Barcode", this.txtPlate_Barcode.Text));
                cmdPlateInsert.Parameters.Add(new OleDbParameter("@plate_Size", this.cmbPlate_Size.SelectionBoxItem.ToString()));
                cmdPlateInsert.Parameters.Add(new OleDbParameter("@in_Use", this.cbPlate_In_Use.IsChecked));
                if (cbPlate_In_Use.IsChecked == true)
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@is_Use_By", Security.employeeID));
                else
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@in_Use_By", DBNull.Value));
                if (cbPlate_In_Use.IsChecked == true)
                {
                    this.txtPlate_Uses_Today.Text = "1";
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@uses_Today", Convert.ToInt32(this.txtPlate_Uses_Today.Text)));
                }
                else if(string.IsNullOrEmpty(this.txtPlate_Uses_Today.Text))
                {
                    this.txtPlate_Uses_Today.Text = "0";
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@uses_Today", Convert.ToInt32(this.txtPlate_Uses_Today.Text)));
                }
                else
                {
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@uses_Today", Convert.ToInt32(this.txtPlate_Uses_Today.Text)));
                }
                if (string.IsNullOrEmpty(this.txtPlate_Uses_Today.Text))
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@total_Uses", DBNull.Value));
                else
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@total_Uses", Convert.ToInt32(this.txtPlate_Uses_Today.Text)));
                cmdPlateInsert.Parameters.Add(new OleDbParameter("@overdue_Cleaning", this.cbPlate_Overdue_Cleaning.IsChecked));
                cmdPlateInsert.Parameters.Add(new OleDbParameter("@location", this.cmbPlate_Location.SelectionBoxItem.ToString()));
                if (string.IsNullOrEmpty(this.txtPlate_Photo.Text))
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@plate_Photo", DBNull.Value));
                else
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@plate_Photo", this.txtPlate_Photo.Text));
                cmdPlateInsert.Parameters.Add(new OleDbParameter("@purchase_Date", this.dpPlate_Purchase_Date.SelectedDate));
                cmdPlateInsert.Parameters.Add(new OleDbParameter("@warranty_Expire", this.dpPlate_Warranty_Expire.SelectedDate));
                if (this.dpPlate_Out_Of_Service.SelectedDate != null)
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@out_Of_Service", this.dpPlate_Out_Of_Service.SelectedDate));
                else
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@out_Of_Service", DBNull.Value));
                if (this.dpPlate_Date_Replaced.SelectedDate != null)
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@date_Replaced", this.dpPlate_Date_Replaced.SelectedDate));
                else
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@date_Replaced", DBNull.Value));
                if (string.IsNullOrEmpty(this.txtPlate_Notes.Text))
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdPlateInsert.Parameters.Add(new OleDbParameter("@notes", this.txtPlate_Notes.Text));

                if (MessageBox.Show("Are you sure you want to add a new plate?", "Add Plate Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdPlateInsert.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblPlates", "tblPlates");
                tblPlatesDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        string sqlPlatesRefresh = "SELECT * FROM tblPlates";
        string tblPlatesRefresh = "tblPlates";
        private void Button_Click_tabPlatesRefresh(object sender, RoutedEventArgs e)
        {
            getDataSet(sqlPlatesRefresh, tblPlatesRefresh);

            if (tblPlatesGrid.Visibility == Visibility.Visible)
            {
                dataGrid_Dates = tblPlatesDataGrid;
                dataGrid_Text = tblPlatesDataGrid;
                dataGrid_Combo = tblPlatesDataGrid;
            }
            else
            {
                dataGrid_Dates = tblPlateCleaningsDataGrid;
                dataGrid_Text = tblPlateCleaningsDataGrid;
                dataGrid_Combo = tblPlateCleaningsDataGrid;
            }

            dataGrid_Dates.DataContext = dataSet.Tables[0];
            dataGrid_Text.DataContext = dataSet.Tables[0];
            dataGrid_Combo.DataContext = dataSet.Tables[0];
        }

        
        //Update info about a plate in tblPlates
        private void Button_Click_btnUpdatePlate(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdPlateUpdate = new OleDbCommand("UPDATE tblPlates SET "
                + "plate_Barcode = @plate_Barcode, "
                + "plate_Size = @plate_Size, "
                + "in_Use = @in_Use, "
                + "in_Use_By = @in_Use_By, "
                + "uses_Today = @uses_Today, "
                + "total_Uses = @total_Uses, "
                + "overdue_Cleaning = @overdue_Cleaning, "
                + "location = @location, "
                + "plate_Photo = @plate_Photo, "
                + "purchase_Date = @purchase_Date, "
                + "warranty_Expire = @warranty_Expire, "
                + "out_Of_Service = @out_Of_Service, "
                + "date_Replaced = @date_Replaced, "
                + "notes = @notes "
                + "WHERE plate_Barcode = @plate_Barcode", dbConnection);
                cmdPlateUpdate.Parameters.Add(new OleDbParameter("@plate_Barcode", this.txtPlate_Barcode.Text));
                cmdPlateUpdate.Parameters.Add(new OleDbParameter("@plate_Size", this.cmbPlate_Size.SelectionBoxItem.ToString()));
                cmdPlateUpdate.Parameters.Add(new OleDbParameter("@in_Use", this.cbPlate_In_Use.IsChecked));
                if (cbPlate_In_Use.IsChecked == true)
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@is_Use_By", Security.employeeID));
                else
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@in_Use_By", DBNull.Value));
                if (cbPlate_In_Use.IsChecked == true)
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@uses_Today", (Convert.ToInt32(this.txtPlate_Uses_Today.Text)) + 1));
                else
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@uses_Today", Convert.ToInt32(this.txtPlate_Uses_Today.Text)));
                if (cbPlate_In_Use.IsChecked == true)
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@total_Uses", (Convert.ToInt32(this.txtPlate_Total_Uses.Text)) + 1));
                else
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@total_Uses", Convert.ToInt32(this.txtPlate_Total_Uses.Text)));
                cmdPlateUpdate.Parameters.Add(new OleDbParameter("@overdue_Cleaning", this.cbPlate_Overdue_Cleaning.IsChecked));
                cmdPlateUpdate.Parameters.Add(new OleDbParameter("@location", this.cmbPlate_Location.SelectionBoxItem.ToString()));
                cmdPlateUpdate.Parameters.Add(new OleDbParameter("@plate_Photo", this.txtPlate_Photo.Text));
                cmdPlateUpdate.Parameters.Add(new OleDbParameter("@purchase_Date", this.dpPlate_Purchase_Date.SelectedDate));
                cmdPlateUpdate.Parameters.Add(new OleDbParameter("@warranty_Expire", this.dpPlate_Warranty_Expire.SelectedDate));
                if (this.dpPlate_Out_Of_Service.SelectedDate != null)
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@out_Of_Service", this.dpPlate_Out_Of_Service.SelectedDate));
                else
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@out_Of_Service", DBNull.Value));
                if (this.dpPlate_Date_Replaced.SelectedDate != null)
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@date_Replaced", this.dpPlate_Date_Replaced.SelectedDate));
                else
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@date_Replaced", DBNull.Value));
                if (string.IsNullOrEmpty(this.txtPlate_Notes.Text))
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdPlateUpdate.Parameters.Add(new OleDbParameter("@notes", this.txtPlate_Notes.Text));
                
                if (MessageBox.Show("Are you sure you want to save?", "Update Plate Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    
                    cmdPlateUpdate.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblPlates", "tblPlates");
                tblPlatesDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
         }

        //Delete a record of a plate from the tblPlates
        private void Button_Click_btnDeletePlate(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdPlateDelete = new OleDbCommand("DELETE FROM tblPlates WHERE plate_Barcode = @plate_Barcode", dbConnection);
                cmdPlateDelete.Parameters.Add(new OleDbParameter("@plate_Barcode", this.txtPlate_Barcode.Text));
                if (MessageBox.Show("Are you sure you want to delete all of the data on this plate?", "Plate Record Deletion", MessageBoxButton.YesNo)==MessageBoxResult.Yes)
                {
                    cmdPlateDelete.ExecuteNonQuery();
                    MessageBox.Show("Delete Successful");
                }
                else
                {
                    this.Close();
                }   
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblPlates", "tblPlates");
                tblPlatesDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        //Alternate between tblPlatesGrid and tblPlateCleaningsGrid
        private void Button_Click_btnViewPlateCleanings(object sender, RoutedEventArgs e)
        {
            if (tblPlatesGrid.Visibility == Visibility.Visible)
            {
                tblPlatesGrid.Visibility = Visibility.Collapsed;
                tblPlateCleaningsGrid.Visibility = Visibility.Visible;
                dkpPlatesQ.Visibility = Visibility.Collapsed;
                dkpPlateCleaningsQ.Visibility = Visibility.Visible;
            }
        }

        private void Button_Click_btnViewPlates(object sender, RoutedEventArgs e)
        {
            if (tblPlatesGrid.Visibility == Visibility.Collapsed)
            {
                tblPlatesGrid.Visibility = Visibility.Visible;
                tblPlateCleaningsGrid.Visibility = Visibility.Collapsed;
                dkpPlatesQ.Visibility = Visibility.Visible;
                dkpPlateCleaningsQ.Visibility = Visibility.Collapsed;
            }
        }

       

        //Insert new plate cleaning record in tblPlateCleanings
        private void Button_Click_btnInsertPlateCleaning(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            { 
                OleDbCommand cmdPlateCleaningInsert = new OleDbCommand("INSERT INTO tblPlateCleanings(plate_Barcode, date_Cleaned, cleaned_By, date_Next_Cleaning, notes) VALUES("
                + "@plate_Barcode, "
                + "@date_Cleaned, "
                + "@cleaned_By, "
                + "@date_Next_Cleaning, "
                + "@notes)", dbConnection);
                cmdPlateCleaningInsert.Parameters.Add(new OleDbParameter("@plate_Barcode", this.txtPlate_Barcode_tblPC.Text));
                cmdPlateCleaningInsert.Parameters.Add(new OleDbParameter("@date_Cleaned", this.dpPlate_Date_Cleaned.SelectedDate));
                cmdPlateCleaningInsert.Parameters.Add(new OleDbParameter("@cleaned_By", Security.employeeID));
                cmdPlateCleaningInsert.Parameters.Add(new OleDbParameter("@date_Next_Cleaning", this.dpPlate_Date_Next_Cleaning.SelectedDate));
                if (string.IsNullOrEmpty(this.txtPlate_Cleaning_Notes.Text))
                    cmdPlateCleaningInsert.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdPlateCleaningInsert.Parameters.Add(new OleDbParameter("@notes", this.txtPlate_Cleaning_Notes.Text));

                if (MessageBox.Show("Are you sure you want to add a new plate cleaning?", "Add Plate Cleaning Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdPlateCleaningInsert.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblPlateCleanings", "tblPlateCleanings");
                tblPlateCleaningsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        //Check to see if user has security clearance
        bool permissionDenied = false;
        //Update plate cleaning record in tblPlateCleanings
        private void Button_Click_btnUpdatePlateCleaning(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            { 
                OleDbCommand cmdPlateCleaningUpdate = new OleDbCommand("UPDATE tblPlateCleanings SET "
                + "plate_Barcode = ?, "
                + "date_Cleaned = ?, "
                + "cleaned_By = ?, "
                + "date_Next_Cleaning = ?, "
                + "notes = ? "
                + "WHERE plate_Cleaning_ID = ?", dbConnection);
                cmdPlateCleaningUpdate.Parameters.Add(new OleDbParameter("@plate_Barcode", this.txtPlate_Barcode_tblPC.Text));
                cmdPlateCleaningUpdate.Parameters.Add(new OleDbParameter("@date_Cleaned", this.dpPlate_Date_Cleaned.SelectedDate));
                if (Security.securityLevel >= 3)
                {
                    cmdPlateCleaningUpdate.Parameters.Add(new OleDbParameter("@cleaned_By", Security.employeeID));
                }
                else if (this.txtPlate_Cleaned_By.Text == Security.employeeID.ToString())
                {
                    cmdPlateCleaningUpdate.Parameters.Add(new OleDbParameter("@cleaned_By", Security.employeeID));
                }
                else
                {
                    permissionDenied = true;
                }
                
                cmdPlateCleaningUpdate.Parameters.Add(new OleDbParameter("@date_Next_Cleaning", this.dpPlate_Date_Next_Cleaning.SelectedDate));
                if (string.IsNullOrEmpty(this.txtPlate_Cleaning_Notes.Text))
                    cmdPlateCleaningUpdate.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdPlateCleaningUpdate.Parameters.Add(new OleDbParameter("@notes", this.txtPlate_Cleaning_Notes.Text));
                cmdPlateCleaningUpdate.Parameters.Add(new OleDbParameter("@plate_Cleaning_ID", this.txtPlate_Cleaning_ID.Text));

                if (permissionDenied == false)
                {
                    if (MessageBox.Show("Are you sure you want to save?", "Update Plate Cleaning Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        cmdPlateCleaningUpdate.ExecuteNonQuery();
                        MessageBox.Show("Save Successful");
                    }
                }
                else
                {
                    MessageBox.Show("You do not have permission to edit this record.\r\nContact an administrator if this record needs to be edited.");
                    permissionDenied = false;
                } 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblPlateCleanings", "tblPlateCleanings");
                tblPlateCleaningsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        //Delete plate cleaning record in tblPlateCleanings
        private void Button_Click_btnDeletePlateCleaning(object sender, RoutedEventArgs e)
        {
            dbConnection.Open(); 
            try
            {
                OleDbCommand cmdPlateCleaningDelete = new OleDbCommand("DELETE FROM tblPlateCleanings WHERE plate_Cleaning_ID = @plate_Cleaning_ID", dbConnection);
                cmdPlateCleaningDelete.Parameters.Add(new OleDbParameter("@plate_Cleaning_ID", Convert.ToInt32(this.txtPlate_Cleaning_ID.Text)));
                if(MessageBox.Show("Are you sure you want to delete all of the data on this plate cleaning?", "Delete Plate Cleaning Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdPlateCleaningDelete.ExecuteNonQuery();
                    MessageBox.Show("Delete Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblPlateCleanings", "tblPlateCleanings");
                tblPlateCleaningsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        
        private void selectedQueryButton(object sender, RoutedEventArgs e)
        {
            if(tabPlates.IsSelected)
            {
                if(dkpPlatesQ.Visibility == Visibility.Visible)
                    Button_Click_btnRunPlatesQuery(sender, e);
                else
                    Button_Click_btnRunPlateCleaningsQuery(sender, e);
            }
            else if(tabCassettes.IsSelected)
            {
                if(dkpCassettesQ.Visibility == Visibility.Visible)
                    Button_Click_btnRunCassettesQuery(sender, e);
                else if(dkpCassetteCleaningsQ.Visibility == Visibility.Visible)
                    Button_Click_btnRunCassetteCleaningsQuery(sender, e);
                else
                    Button_Click_btnRunCassetteServicingsQuery(sender, e);
            }
        }
        
        //Query functions
        string sqlSelect_Dates;
        string tblName_Dates;
        DataGrid dataGrid_Dates;
        string parameterName_Dates;
        object parameterValue_Dates;
        private void Button_Click_btnDateQuerySubmit(object sender, RoutedEventArgs e)
        {
            selectedQueryButton(sender, e);
            dbConnection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            OleDbCommand cmdSelect = new OleDbCommand(sqlSelect_Dates, dbConnection);
            adapter.SelectCommand = cmdSelect;
            cmdSelect.Parameters.Add(parameterName_Dates, OleDbType.Date).Value = parameterValue_Dates;
            DataSet results = new DataSet();
            adapter.Fill(results, tblName_Dates);
            dataGrid_Dates.DataContext = results.Tables[0];
            if (dataGrid_Dates == queryResultsDataGrid)
                queryResultsGrid.Visibility = Visibility.Visible;
            dbConnection.Close();
            dateInputBoxGrid.Visibility = Visibility.Collapsed;
        }


        string sqlSelect_Text;
        string tblName_Text;
        DataGrid dataGrid_Text;
        string parameterName_Text;
        object parameterValue_Text;
        private void Button_Click_btnTextQuerySubmit(object sender, RoutedEventArgs e)
        {
            selectedQueryButton(sender, e);
            dbConnection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            OleDbCommand cmdSelect = new OleDbCommand(sqlSelect_Text, dbConnection);
            adapter.SelectCommand = cmdSelect;
            cmdSelect.Parameters.Add(new OleDbParameter(parameterName_Text, parameterValue_Text));
            DataSet results = new DataSet();
            adapter.Fill(results, tblName_Text);
            dataGrid_Text.DataContext = results.Tables[0];
            if(dataGrid_Text == queryResultsDataGrid)
                queryResultsGrid.Visibility = Visibility.Visible;
            dbConnection.Close();
            textInputBoxGrid.Visibility = Visibility.Collapsed;
        }

        string sqlSelect_Combo;
        string tblName_Combo;
        DataGrid dataGrid_Combo;
        string parameterName_Combo;
        object parameterValue_Combo;
        private void Button_Click_btnComboQuerySubmit(object sender, RoutedEventArgs e)
        {
            selectedQueryButton(sender, e);
            dbConnection.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            OleDbCommand cmdSelect = new OleDbCommand(sqlSelect_Combo, dbConnection);
            adapter.SelectCommand = cmdSelect;
            cmdSelect.Parameters.Add(new OleDbParameter(parameterName_Combo, parameterValue_Combo));
            DataSet results = new DataSet();
            adapter.Fill(results, tblName_Combo);
            dataGrid_Combo.DataContext = results.Tables[0];
            if (dataGrid_Combo == queryResultsDataGrid)
                queryResultsGrid.Visibility = Visibility.Visible;
            dbConnection.Close();
            comboInputBoxGrid.Visibility = Visibility.Collapsed;
            cmbInputCombobox.Items.Clear();
        }

        private void Button_Click_closeQueryResults(object sender, RoutedEventArgs e)
        {
            queryResultsGrid.Visibility = Visibility.Collapsed;
        }

        OleDbDataAdapter da;
        OleDbCommand cmd;
        DataSet results;
        private void Button_Click_btnRunPlatesQuery(object sender, RoutedEventArgs e)
        {
            switch(cmbPlatesQuery.Text)
            { 
                case "Plates by barcode":
                    txbTextInputInstructions.Text = "Please enter a number of weeks";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblPlates";
                    sqlSelect_Text = "SELECT * FROM tblPlates WHERE plate_Barcode = ?";
                    dataGrid_Text = tblPlatesDataGrid;
                    parameterName_Text = "@plate_Barcode";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    break;

                case "New Plates":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblPlates WHERE total_Uses = 0", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblPlates");
                    tblPlatesDataGrid.DataContext = results.Tables[0];
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    dbConnection.Close();
                    break;
               
                case "Plates due for cleaning":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblPlates WHERE overdue_Cleaning = Yes", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblPlates");
                    tblPlatesDataGrid.DataContext = results.Tables[0];
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    dbConnection.Close();
                    break;

                case "Plates needing replacement":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblPlates WHERE out_Of_Service <> Null AND date_Replaced Is Null", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblPlates");
                    tblPlatesDataGrid.DataContext = results.Tables[0];
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    dbConnection.Close();
                    break;

                case "Plates bought by date":
                    txbDateInputInstructions.Text = "Please enter a date";
                    dateInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Dates = "tblPlates";
                    sqlSelect_Dates = "SELECT * FROM tblPlates WHERE purchase_Date = ?";
                    dataGrid_Dates = tblPlatesDataGrid;
                    parameterName_Dates = "@purchase_Date";
                    parameterValue_Dates = this.dpInputDatePicker.SelectedDate;
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    break;

                case "Plates bought in past given weeks":
                    txbTextInputInstructions.Text = "Please enter a number of weeks";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblPlates";
                    sqlSelect_Text = "SELECT * FROM tblPlates WHERE purchase_Date >= now - (?*7)";
                    dataGrid_Text = tblPlatesDataGrid;
                    parameterName_Text = "@numberOfWeeks";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    break;

                case "Plates bought in past given months":
                    txbTextInputInstructions.Text = "Please enter a number of months";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblPlates";
                    sqlSelect_Text = "SELECT * FROM tblPlates WHERE purchase_Date >= now - (?*30)";
                    dataGrid_Text = tblPlatesDataGrid;
                    parameterName_Text = "@numberOfWeeks";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    break;

                case "Expired plate warranties":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblPlates WHERE warranty_Expire <= now", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblPlates");
                    tblPlatesDataGrid.DataContext = results.Tables[0];
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    dbConnection.Close();
                    break;

                case "Plate warranties expiring in given number of weeks":
                    txbTextInputInstructions.Text = "Please enter a number of weeks";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblPlates";
                    sqlSelect_Text = "SELECT * FROM tblPlates WHERE warranty_Expire <= now + (?*7)";
                    dataGrid_Text = tblPlatesDataGrid;
                    parameterName_Text = "@numberOfWeeks";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    break;

                case "Plates by location":
                    fillComboBox("SELECT * FROM tblComboLocation", cmbInputCombobox);
                    txbComboInputInstructions.Text = "Please select a location";
                    comboInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Combo = "tblPlates";
                    sqlSelect_Combo = "SELECT * FROM tblPlates WHERE location = ?";
                    dataGrid_Combo = tblPlatesDataGrid;
                    parameterName_Combo = "@location";
                    parameterValue_Combo = this.cmbInputCombobox.Text;
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlates";
                    break;

                case "Plate by who is currently using it":
                    txbTextInputInstructions.Text = "Please enter a username";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblPlates";
                    sqlSelect_Text = "SELECT tblPlates.plate_Barcode, "
                                    + "tblPlates.plate_Size, "
                                    + "tblUsers.first_Name, "
                                    + "tblUsers.last_Name, "
                                    + "tblPlates.uses_Today, "
                                    + "tblPlates.total_Uses, "
                                    + "tblPlates.overdue_Cleaning, "
                                    + "tblPlates.plate_Photo, "
                                    + "tblPlates.warranty_Expire, "
                                    + "tblPlates.notes "
                                    + "FROM tblPlates INNER JOIN tblUsers ON tblPlates.in_Use_By = tblUsers.employee_ID "
                                    + "WHERE tblUsers.username = ?";
                    dataGrid_Text = queryResultsDataGrid;
                    parameterName_Text = "@username";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlPlatesRefresh = "SELECT * FROM tblPlates";
                    tblPlatesRefresh = "tblPlateCleanings";
                    break;
            }
        }


        private void Button_Click_btnRunPlateCleaningsQuery(object sender, RoutedEventArgs e)
        {
            switch (cmbPlateCleaningsQuery.Text)
            {
                case "Plates due for cleaning in the next week":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblPlateCleanings WHERE tblPlateCleanings.date_Next_Cleaning <= now + 7", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblPlateCleanings");
                    tblPlateCleaningsDataGrid.DataContext = results.Tables[0];
                    sqlPlatesRefresh = "SELECT * FROM tblPlateCleanings";
                    tblPlatesRefresh = "tblPlateCleanings";
                    dbConnection.Close();
                    break;
                
                case "Plate cleanings by date":
                    txbDateInputInstructions.Text = "Please enter a date";
                    dateInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Dates = "tblPlateCleanings";
                    sqlSelect_Dates = "SELECT tblPlateCleanings.plate_Cleaning_ID, "
                                            + "tblPlateCleanings.plate_Barcode, "
                                            + "tblPlateCleanings.date_Cleaned, "
                                            + "tblUsers.first_Name, "
                                            + "tblUsers.last_Name, "
                                            + "tblPlateCleanings.date_Next_Cleaning, "
                                            + "tblPlateCleanings.notes "
                                            + "FROM tblPlateCleanings INNER JOIN tblUsers ON tblPlateCleanings.cleaned_By = tblUsers.employee_ID "
                                            + "WHERE tblPlateCleanings.date_Cleaned = ?";
                    dataGrid_Dates = queryResultsDataGrid;
                    parameterName_Dates = "@date_Cleaned";
                    parameterValue_Dates = this.dpInputDatePicker.SelectedDate;
                    sqlPlatesRefresh = "SELECT * FROM tblPlateCleanings";
                    tblPlatesRefresh = "tblPlateCleanings";
                    break;

                case "Plate cleanings by plate cleaning ID":
                    txbTextInputInstructions.Text = "Please enter a plate cleaning ID";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblPlateCleanings";
                    sqlSelect_Text = "SELECT tblPlateCleanings.plate_Cleaning_ID, "
                                            + "tblPlateCleanings.plate_Barcode, "
                                            + "tblPlateCleanings.date_Cleaned, "
                                            + "tblUsers.first_Name, "
                                            + "tblUsers.last_Name, "
                                            + "tblPlateCleanings.date_Next_Cleaning, "
                                            + "tblPlateCleanings.notes "
                                            + "FROM tblPlateCleanings INNER JOIN tblUsers ON tblPlateCleanings.cleaned_By = tblUsers.employee_ID "
                                            + "WHERE plate_Cleaning_ID = ?";
                    dataGrid_Text = queryResultsDataGrid;
                    parameterName_Text = "@plate_Cleaning_ID";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlPlatesRefresh = "SELECT * FROM tblPlateCleanings";
                    tblPlatesRefresh = "tblPlateCleanings";
                    break;

                case "Plate cleanings by plate barcode":
                    txbTextInputInstructions.Text = "Please enter a plate barcode";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblPlateCleanings";
                    sqlSelect_Text = "SELECT * FROM tblPlateCleanings WHERE plate_Barcode = ?";
                    dataGrid_Text =tblPlateCleaningsDataGrid;
                    parameterName_Text = "@plate_Barcode";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlPlatesRefresh = "SELECT * FROM tblPlateCleanings";
                    tblPlatesRefresh = "tblPlateCleanings";
                    break;
            }
        }

        private void Button_Click_btnRunCassettesQuery(object sender, RoutedEventArgs e)
        {
            switch (cmbCassettesQuery.Text)
            {
                case "Cassettes due for cleaning":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblCassettes WHERE overdue_Cleaning = YES;", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblCassettes");
                    tblCassettesDataGrid.DataContext = results.Tables[0];
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    dbConnection.Close();
                    break;

                case "New Cassettes":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblCassettes WHERE total_Uses = 0;", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblCassettes");
                    tblCassettesDataGrid.DataContext = results.Tables[0];
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    dbConnection.Close();
                    break;

                case "Cassettes needing replacement":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblCassettes WHERE out_Of_Service <> NULL;", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblCassettes");
                    tblCassettesDataGrid.DataContext = results.Tables[0];
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    dbConnection.Close();
                    break;

                case "Cassettes bought by date":
                    txbDateInputInstructions.Text = "Please enter a date";
                    dateInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Dates = "tblCassettes";
                    sqlSelect_Dates = "SELECT * FROM tblCassettes WHERE purchase_Date = ?";
                    dataGrid_Dates = tblCassettesDataGrid;
                    parameterName_Dates = "@purchase_Date";
                    parameterValue_Dates = this.dpInputDatePicker.SelectedDate;
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    break;

                case "Cassettes bought in past given weeks":
                    txbTextInputInstructions.Text = "Please enter a number of weeks";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblCassettes";
                    sqlSelect_Text = "SELECT * FROM tblCassettes WHERE purchase_Date >= now - (?*7)";
                    dataGrid_Text = tblCassettesDataGrid;
                    parameterName_Text = "@numberOfWeeks";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    break;

                case "Cassettes bought in past given months":
                    txbTextInputInstructions.Text = "Please enter a number of months";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblCassettes";
                    sqlSelect_Text = "SELECT * FROM tblCassettes WHERE purchase_Date >= now - (?*30)";
                    dataGrid_Text = tblCassettesDataGrid;
                    parameterName_Text = "@numberOfWeeks";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    break;

                case "Expired cassette warranties":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblCassettes WHERE warranty_Expire <= now", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblCassettes");
                    tblCassettesDataGrid.DataContext = results.Tables[0];
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    dbConnection.Close();
                    break;

                case "Cassette warranties expiring in given number of weeks":
                    txbTextInputInstructions.Text = "Please enter a number of weeks";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblCassettes";
                    sqlSelect_Text = "SELECT * FROM tblCassettes WHERE warranty_Expire =< now + (?*7)";
                    dataGrid_Text = tblCassettesDataGrid;
                    parameterName_Text = "@numberOfWeeks";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    break;

                case "Cassettes by location":
                    fillComboBox("SELECT * FROM tblComboLocation", cmbInputCombobox);
                    txbComboInputInstructions.Text = "Please select a location";
                    comboInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Combo = "tblCassettes";
                    sqlSelect_Combo = "SELECT * FROM tblCassettes WHERE location = ?";
                    parameterName_Combo = "@location";
                    dataGrid_Combo = tblCassettesDataGrid;
                    parameterValue_Combo = this.cmbInputCombobox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    break;

                case "Cassette by who is currently using it":
                    txbTextInputInstructions.Text = "Please enter a username";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblCassettes";
                    sqlSelect_Text = "SELECT tblCassettes.cassette_Barcode, "
                                    + "tblCassettes.cassette_Size, "
                                    + "tblUsers.first_Name, "
                                    + "tblUsers.last_Name, "
                                    + "tblCassettes.uses_Today, "
                                    + "tblCassettes.total_Uses, "
                                    + "tblCassettes.overdue_Cleaning, "
                                    + "tblCassettes.cassette_Photo, "
                                    + "tblCassettes.warranty_Expire, "
                                    + "tblCassettes.notes "
                                    + "FROM tblCassettes INNER JOIN tblUsers ON tblCassettes.in_Use_By = tblUsers.employee_ID "
                                    + "WHERE tblUsers.username = ?";
                    dataGrid_Text = queryResultsDataGrid;
                    parameterName_Text = "@username";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassetteCleanings";
                    break;

                case "Cassettes sent out for service":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblCassettes WHERE sent_For_Service <> NULL;", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblCassettes");
                    tblCassettesDataGrid.DataContext = results.Tables[0];
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    dbConnection.Close();
                    break;

                case "Cassettes sent out to be serviced by date":
                    txbDateInputInstructions.Text = "Please enter a date";
                    dateInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Dates = "tblCassettes";
                    sqlSelect_Dates = "SELECT * FROM tblCassettes WHERE date_Last_Serviced = ?;";
                    dataGrid_Dates = tblCassettesDataGrid;
                    parameterName_Dates = "@date_Last_Serviced";
                    parameterValue_Dates = this.dpInputDatePicker.SelectedDate;
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    break;

                case "Cassettes that need to be sent out for service":
                     dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblCassettes WHERE needs_Service = YES;", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblCassettes");
                    tblCassettesDataGrid.DataContext = results.Tables[0];
                    sqlCassettesRefresh = "SELECT * FROM tblCassettes";
                    tblCassettesRefresh = "tblCassettes";
                    dbConnection.Close();
                    break;
            }
        }

        private void Button_Click_btnRunCassetteCleaningsQuery(object sender, RoutedEventArgs e)
        {
            switch (cmbCassetteCleaningsQuery.Text)
            {
                case "Cassette cleanings by cassette cleaning ID":
                    txbTextInputInstructions.Text = "Please enter a cassette cleaning ID";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblCassetteCleanings";
                    sqlSelect_Text = "SELECT tblCassetteCleanings.cassette_Cleaning_ID, "
                                            + "tblCassetteCleanings.cassette_Barcode, "
                                            + "tblCassetteCleanings.date_Cleaned, "
                                            + "tblUsers.first_Name, "
                                            + "tblUsers.last_Name, "
                                            + "tblCassetteCleanings.date_Next_Cleaning, "
                                            + "tblCassetteCleanings.notes "
                                            + "FROM tblCassetteCleanings INNER JOIN tblUsers ON tblCassetteCleanings.cleaned_By = tblUsers.employee_ID "
                                            + "WHERE cassette_Cleaning_ID = ?";
                    dataGrid_Text = queryResultsDataGrid;
                    parameterName_Text = "@cassette_Cleaning_ID";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblcassetteCleanings";
                    tblCassettesRefresh = "tblcassetteCleanings";
                    break;

                case "Cassette cleanings by cassette barcode":
                    txbTextInputInstructions.Text = "Please enter a cassette barcode";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblCassetteCleanings";
                    sqlSelect_Text = "SELECT * FROM tblCassetteCleanings WHERE cassette_Barcode = ?";
                    dataGrid_Text = tblCassetteCleaningsDataGrid;
                    parameterName_Text = "@cassette_Barcode";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblCassetteCleanings";
                    tblCassettesRefresh = "tblCassetteCleanings";
                    break;

                case "Cassette cleanings by date":
                    txbDateInputInstructions.Text = "Please enter a date";
                    dateInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Dates = "tblCassetteCleanings";
                    sqlSelect_Dates = "SELECT tblCassetteCleanings.cassette_Cleaning_ID, "
                                            + "tblCassetteCleanings.cassette_Barcode, "
                                            + "tblCassetteCleanings.date_Cleaned, "
                                            + "tblUsers.first_Name, "
                                            + "tblUsers.last_Name, "
                                            + "tblCassetteCleanings.date_Next_Cleaning, "
                                            + "tblCassetteCleanings.notes "
                                            + "FROM tblCassetteCleanings INNER JOIN tblUsers ON tblCassetteCleanings.cleaned_By = tblUsers.employee_ID "
                                            + "WHERE tblCassetteCleanings.date_Cleaned = ?";
                    dataGrid_Dates = queryResultsDataGrid;
                    parameterName_Dates = "@date_Cleaned";
                    parameterValue_Dates = this.dpInputDatePicker.SelectedDate;
                    sqlCassettesRefresh = "SELECT * FROM tblCassetteCleanings";
                    tblCassettesRefresh = "tblCassetteCleanings";
                    break;

                case "Cassettes due for cleaning in the next week":
                    dbConnection.Open();
                    da = new OleDbDataAdapter();
                    cmd = new OleDbCommand("SELECT * FROM tblCassetteCleanings WHERE tblCassetteCleanings.date_Next_Cleaning <= now + 7", dbConnection);
                    da.SelectCommand = cmd;
                    results = new DataSet();
                    da.Fill(results, "tblCassetteCleanings");
                    tblCassetteCleaningsDataGrid.DataContext = results.Tables[0];
                    sqlCassettesRefresh = "SELECT * FROM tblCassetteCleanings";
                    tblCassettesRefresh = "tblCassetteCleanings";
                    dbConnection.Close();
                    break;
            }
        }

        private void Button_Click_btnRunCassetteServicingsQuery(object sender, RoutedEventArgs e)
        {
            switch (cmbCassetteServicingsQuery.Text)
            {
                case "Cassette servicings by cassette servicing ID":
                    txbTextInputInstructions.Text = "Please enter a cassette servicing ID";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblCassetteServicings";
                    sqlSelect_Text = "SELECT * FROM tblCassetteServicings WHERE cassette_Servicing_ID = ?";
                    dataGrid_Text = tblCassetteServicingsDataGrid;
                    parameterName_Text = "@cassette_Servicing_ID";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblCassetteServicings";
                    tblCassettesRefresh = "tblCassetteServicings";
                    break;

                case "Cassette servicings by barcode":
                    txbTextInputInstructions.Text = "Please enter a cassette barcode";
                    textInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Text = "tblCassetteServicings";
                    sqlSelect_Text = "SELECT * FROM tblCassetteServicings WHERE cassette_Barcode = ?";
                    dataGrid_Text = tblCassetteServicingsDataGrid;
                    parameterName_Text = "@cassette_Barcode";
                    parameterValue_Text = this.txtInputTextBox.Text;
                    sqlCassettesRefresh = "SELECT * FROM tblCassetteServicings";
                    tblCassettesRefresh = "tblCassetteServicings";
                    break;

                case "Cassette servicings by date servicing needed":
                    txbDateInputInstructions.Text = "Please enter a date";
                    dateInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Dates = "tblCassetteServicings";
                    sqlSelect_Dates = "SELECT * FROM tblCassetteServicings WHERE date_Needs_Servicing = ?";
                    dataGrid_Dates = tblCassetteServicingsDataGrid;
                    parameterName_Dates = "@date_Needs_Servicing";
                    parameterValue_Dates = this.dpInputDatePicker.SelectedDate;
                    sqlCassettesRefresh = "SELECT * FROM tblCassetteServicings";
                    tblCassettesRefresh = "tblCassetteServicings";
                    break;

                case "Cassette servicings by date back in service":
                    txbDateInputInstructions.Text = "Please enter a date";
                    dateInputBoxGrid.Visibility = Visibility.Visible;
                    tblName_Dates = "tblCassetteServicings";
                    sqlSelect_Dates = "SELECT * FROM tblCassetteServicings WHERE date_Back_In_Service = ?";
                    dataGrid_Dates = tblCassetteServicingsDataGrid;
                    parameterName_Dates = "@date_Back_In_Service";
                    parameterValue_Dates = this.dpInputDatePicker.SelectedDate;
                    sqlCassettesRefresh = "SELECT * FROM tblCassetteServicings";
                    tblCassettesRefresh = "tblCassetteServicings";
                    break;
            }
        }

        string sqlCassettesRefresh;
        string tblCassettesRefresh;
        private void Button_Click_tabCassettesRefresh(object sender, RoutedEventArgs e)
        {
            getDataSet(sqlCassettesRefresh, tblCassettesRefresh);

            if (tblCassettesGrid.Visibility == Visibility.Visible)
            {
                dataGrid_Dates = tblCassettesDataGrid;
                dataGrid_Text = tblCassettesDataGrid;
                dataGrid_Combo = tblCassettesDataGrid;
            }
            else if(tblCassetteCleaningsGrid.Visibility == Visibility.Visible)
            {
                dataGrid_Dates = tblCassetteCleaningsDataGrid;
                dataGrid_Text = tblCassetteCleaningsDataGrid;
                dataGrid_Combo = tblCassetteCleaningsDataGrid;
            }
            else
            {
                dataGrid_Dates = tblCassetteServicingsDataGrid;
                dataGrid_Text = tblCassetteServicingsDataGrid;
                dataGrid_Combo = tblCassetteServicingsDataGrid;
            }

            dataGrid_Dates.DataContext = dataSet.Tables[0];
            dataGrid_Text.DataContext = dataSet.Tables[0];
            dataGrid_Combo.DataContext = dataSet.Tables[0];
        }

        private void Button_Click_btnViewCassettes(object sender, RoutedEventArgs e)
        {
            if (tblCassettesGrid.Visibility == Visibility.Collapsed)
            {
                tblCassetteCleaningsGrid.Visibility = Visibility.Collapsed;
                tblCassetteServicingsGrid.Visibility = Visibility.Collapsed;
                tblCassettesGrid.Visibility = Visibility.Visible;
                dkpCassetteCleaningsQ.Visibility = Visibility.Collapsed;
                dkpCassetteServicingsQ.Visibility = Visibility.Collapsed;
                dkpCassettesQ.Visibility = Visibility.Visible;
            }
        }
        
        private void Button_Click_btnViewCassetteCleanings(object sender, RoutedEventArgs e)
        {
            if (tblCassetteCleaningsGrid.Visibility == Visibility.Collapsed)
            {
                tblCassettesGrid.Visibility = Visibility.Collapsed;
                tblCassetteServicingsGrid.Visibility = Visibility.Collapsed;
                tblCassetteCleaningsGrid.Visibility = Visibility.Visible;
                dkpCassettesQ.Visibility = Visibility.Collapsed;
                dkpCassetteServicingsQ.Visibility = Visibility.Collapsed;
                dkpCassetteCleaningsQ.Visibility = Visibility.Visible;
            }
        }

        private void Button_Click_btnViewCassetteServicings(object sender, RoutedEventArgs e)
        {
            if (tblCassetteServicingsGrid.Visibility == Visibility.Collapsed)
            {
                tblCassettesGrid.Visibility = Visibility.Collapsed;
                tblCassetteCleaningsGrid.Visibility = Visibility.Collapsed;
                dkpCassettesQ.Visibility = Visibility.Collapsed;
                dkpCassetteCleaningsQ.Visibility = Visibility.Collapsed;
                tblCassetteServicingsGrid.Visibility = Visibility.Visible;
                dkpCassetteServicingsQ.Visibility = Visibility.Visible;
            }
        }
        

        //Insert new Cassette record
        private void Button_Click_btnInsertCassette(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteInsert = new OleDbCommand("INSERT INTO tblCassettes VALUES("
                + "@cassette_Barcode, "
                + "@cassette_Size, "
                + "@in_Use, "
                + "@in_Use_By, "
                + "@uses_Today, "
                + "@total_Uses, "
                + "@overdue_Cleaning, "
                + "@location, "
                + "@cassette_Photo, "
                + "@needs_Service, "
                + "@sent_For_Service, "
                + "@purchase_Date, "
                + "@warranty_Expire, "
                + "@out_Of_Service, "
                + "@date_Replaced, "
                + "@notes)", dbConnection);
                cmdCassetteInsert.Parameters.Add(new OleDbParameter("@cassette_Barcode", this.txtCassette_Barcode.Text));
                cmdCassetteInsert.Parameters.Add(new OleDbParameter("@cassette_Size", this.cmbCassette_Size.SelectionBoxItem.ToString()));
                cmdCassetteInsert.Parameters.Add(new OleDbParameter("@in_Use", this.cbCassette_In_Use.IsChecked));
                if (cbCassette_In_Use.IsChecked == true)
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@is_Use_By", Security.employeeID));
                else
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@in_Use_By", DBNull.Value));
                if (cbCassette_In_Use.IsChecked == true)
                {
                    this.txtCassette_Uses_Today.Text = "1";
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@uses_Today", Convert.ToInt32(this.txtCassette_Uses_Today.Text)));
                }
                else if (string.IsNullOrEmpty(this.txtCassette_Uses_Today.Text))
                {
                    this.txtCassette_Uses_Today.Text = "0";
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@uses_Today", Convert.ToInt32(this.txtCassette_Uses_Today.Text)));
                }
                else
                {
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@uses_Today", Convert.ToInt32(this.txtCassette_Uses_Today.Text)));
                }
                if (string.IsNullOrEmpty(this.txtCassette_Uses_Today.Text))
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@total_Uses", DBNull.Value));
                else
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@total_Uses", Convert.ToInt32(this.txtCassette_Uses_Today.Text)));
                cmdCassetteInsert.Parameters.Add(new OleDbParameter("@overdue_Cleaning", this.cbCassette_Overdue_Cleaning.IsChecked));
                cmdCassetteInsert.Parameters.Add(new OleDbParameter("@location", this.cmbCassette_Location.SelectionBoxItem.ToString()));
                if (string.IsNullOrEmpty(this.txtCassette_Photo.Text))
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@cassette_Photo", DBNull.Value));
                else
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@cassette_Photo", this.txtCassette_Photo.Text));
                
                cmdCassetteInsert.Parameters.Add(new OleDbParameter("@needs_Service", this.cbCassette_Needs_Service.IsChecked));
                
                if(this.dpCassette_Date_Last_Serviced.SelectedDate != null)
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@sent_For_Service", this.dpCassette_Date_Last_Serviced.SelectedDate));
                else
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@sent_For_Service", DBNull.Value));
                
                cmdCassetteInsert.Parameters.Add(new OleDbParameter("@purchase_Date", this.dpCassette_Purchase_Date.SelectedDate));
                
                cmdCassetteInsert.Parameters.Add(new OleDbParameter("@warranty_Expire", this.dpCassette_Warranty_Expire.SelectedDate));
                
                if (this.dpCassette_Out_Of_Service.SelectedDate != null)
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@out_Of_Service", this.dpCassette_Out_Of_Service.SelectedDate));
                else
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@out_Of_Service", DBNull.Value));
                if (this.dpCassette_Date_Replaced.SelectedDate != null)
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@date_Replaced", this.dpCassette_Date_Replaced.SelectedDate));
                else
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@date_Replaced", DBNull.Value));
                if (string.IsNullOrEmpty(this.txtCassette_Notes.Text))
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdCassetteInsert.Parameters.Add(new OleDbParameter("@notes", this.txtCassette_Notes.Text));

                if (MessageBox.Show("Are you sure you want to add a new cassette?", "Add Plate Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdCassetteInsert.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassettes", "tblCassettes");
                tblCassettesDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void Button_Click_btnUpdateCassette(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteUpdate = new OleDbCommand("UPDATE tblCassettes SET "
                + "cassette_Barcode = @cassette_Barcode, "
                + "cassette_Size = @cassette_Size, "
                + "in_Use = @in_Use, "
                + "in_Use_By = @in_Use_By, "
                + "uses_Today = @uses_Today, "
                + "total_Uses = @total_Uses, "
                + "overdue_Cleaning = @overdue_Cleaning, "
                + "location = @location, "
                + "cassette_Photo = @cassette_Photo, "
                + "needs_Service = @needs_Service, "
                + "sent_For_Service = @sent_For_Service, "
                + "purchase_Date = @purchase_Date, "
                + "warranty_Expire = @warranty_Expire, "
                + "out_Of_Service = @out_Of_Service, "
                + "date_Replaced = @date_Replaced, "
                + "notes = @notes "
                + "WHERE cassette_Barcode = @cassette_Barcode", dbConnection);
                cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@cassette_Barcode", this.txtCassette_Barcode.Text));
                cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@cassette_Size", this.cmbCassette_Size.SelectionBoxItem.ToString()));
                cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@in_Use", this.cbCassette_In_Use.IsChecked));
                if (cbCassette_In_Use.IsChecked == true)
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@is_Use_By", Security.employeeID));
                else
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@in_Use_By", DBNull.Value));
                if (cbCassette_In_Use.IsChecked == true)
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@uses_Today", (Convert.ToInt32(this.txtCassette_Uses_Today.Text)) + 1));
                else
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@uses_Today", Convert.ToInt32(this.txtCassette_Uses_Today.Text)));
                if (cbCassette_In_Use.IsChecked == true)
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@total_Uses", (Convert.ToInt32(this.txtCassette_Total_Uses.Text)) + 1));
                else
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@total_Uses", Convert.ToInt32(this.txtCassette_Total_Uses.Text)));
                cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@overdue_Cleaning", this.cbCassette_Overdue_Cleaning.IsChecked));
                cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@location", this.cmbCassette_Location.SelectionBoxItem.ToString()));
                if (string.IsNullOrEmpty(this.txtCassette_Photo.Text))
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@cassette_Photo", DBNull.Value));
                else
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@cassette_Photo", this.txtCassette_Photo.Text));

                cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@needs_Service", this.cbCassette_Needs_Service.IsChecked));

                if (this.dpCassette_Date_Last_Serviced.SelectedDate != null)
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@sent_For_Service", this.dpCassette_Date_Last_Serviced.SelectedDate));
                else
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@sent_For_Service", DBNull.Value));

                cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@purchase_Date", this.dpCassette_Purchase_Date.SelectedDate));

                cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@warranty_Expire", this.dpCassette_Warranty_Expire.SelectedDate));

                if (this.dpCassette_Out_Of_Service.SelectedDate != null)
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@out_Of_Service", this.dpCassette_Out_Of_Service.SelectedDate));
                else
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@out_Of_Service", DBNull.Value));
                if (this.dpCassette_Date_Replaced.SelectedDate != null)
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@date_Replaced", this.dpCassette_Date_Replaced.SelectedDate));
                else
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@date_Replaced", DBNull.Value));
                if (string.IsNullOrEmpty(this.txtCassette_Notes.Text))
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdCassetteUpdate.Parameters.Add(new OleDbParameter("@notes", this.txtCassette_Notes.Text));

                if (MessageBox.Show("Are you sure you want to save?", "Update Cassette Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdCassetteUpdate.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassettes", "tblCassettes");
                tblCassettesDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void Button_Click_btnDeleteCassette(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteDelete = new OleDbCommand("DELETE FROM tblCassettes WHERE cassette_Barcode = @cassette_Barcode", dbConnection);
                cmdCassetteDelete.Parameters.Add(new OleDbParameter("@cassette_Barcode", this.txtCassette_Barcode.Text));
                if (MessageBox.Show("Are you sure you want to delete all of the data on this cassette?", "Delete Cassette Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdCassetteDelete.ExecuteNonQuery();
                    MessageBox.Show("Delete Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassettes", "tblCassettes");
                tblCassettesDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void Button_Click_btnInsertCassetteCleaning(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteCleaningInsert = new OleDbCommand("INSERT INTO tblCassetteCleanings(cassette_Barcode, date_Cleaned, cleaned_By, date_Next_Cleaning, notes) VALUES("
                + "@cassette_Barcode, "
                + "@date_Cleaned, "
                + "@cleaned_By, "
                + "@date_Next_Cleaning, "
                + "@notes)", dbConnection);
                cmdCassetteCleaningInsert.Parameters.Add(new OleDbParameter("@cassette_Barcode", this.txtCassette_Barcode_tblCC.Text));
                cmdCassetteCleaningInsert.Parameters.Add(new OleDbParameter("@date_Cleaned", this.dpCassette_Date_Cleaned.SelectedDate));
                cmdCassetteCleaningInsert.Parameters.Add(new OleDbParameter("@cleaned_By", Security.employeeID));
                cmdCassetteCleaningInsert.Parameters.Add(new OleDbParameter("@date_Next_Cleaning", this.dpCassette_Date_Next_Cleaning.SelectedDate));
                if (string.IsNullOrEmpty(this.txtCassette_Cleaning_Notes.Text))
                    cmdCassetteCleaningInsert.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdCassetteCleaningInsert.Parameters.Add(new OleDbParameter("@notes", this.txtCassette_Cleaning_Notes.Text));

                if (MessageBox.Show("Are you sure you want to add a new cassette cleaning?", "Add Cassette Cleaning Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdCassetteCleaningInsert.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassetteCleanings", "tblCassetteCleanings");
                tblCassetteCleaningsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void Button_Click_btnUpdateCassetteCleaning(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteCleaningUpdate = new OleDbCommand("UPDATE tblCassetteCleanings SET "
                + "cassette_Barcode = ?, "
                + "date_Cleaned = ?, "
                + "cleaned_By = ?, "
                + "date_Next_Cleaning = ?, "
                + "notes = ? "
                + "WHERE cassette_Cleaning_ID = ?", dbConnection);
                cmdCassetteCleaningUpdate.Parameters.Add(new OleDbParameter("@cassette_Barcode", this.txtCassette_Barcode_tblCC.Text));
                cmdCassetteCleaningUpdate.Parameters.Add(new OleDbParameter("@date_Cleaned", this.dpCassette_Date_Cleaned.SelectedDate));
                if (Security.securityLevel >= 3)
                {
                    cmdCassetteCleaningUpdate.Parameters.Add(new OleDbParameter("@cleaned_By", Security.employeeID));
                }
                else if (this.txtCassette_Cleaned_By.Text == Security.employeeID.ToString())
                {
                    cmdCassetteCleaningUpdate.Parameters.Add(new OleDbParameter("@cleaned_By", Security.employeeID));
                }
                else
                {
                    permissionDenied = true;
                }
                cmdCassetteCleaningUpdate.Parameters.Add(new OleDbParameter("@date_Next_Cleaning", this.dpCassette_Date_Next_Cleaning.SelectedDate));
                if (string.IsNullOrEmpty(this.txtCassette_Cleaning_Notes.Text))
                    cmdCassetteCleaningUpdate.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdCassetteCleaningUpdate.Parameters.Add(new OleDbParameter("@notes", this.txtCassette_Cleaning_Notes.Text));
                cmdCassetteCleaningUpdate.Parameters.Add(new OleDbParameter("@cassette_Cleaning_ID", Convert.ToInt32(this.txtCassette_Cleaning_ID.Text)));

                if (permissionDenied == false)
                {
                    if (MessageBox.Show("Are you sure you want to save?", "Update Cassette Cleaning Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        cmdCassetteCleaningUpdate.ExecuteNonQuery();
                        MessageBox.Show("Save Successful");
                    }

                }
                else
                {
                    MessageBox.Show("You do not have permission to edit this record.\r\nContact an administrator if this record needs to be edited.");
                    permissionDenied = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassetteCleanings", "tblCassetteCleanings");
                tblCassetteCleaningsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void Button_Click_btnDeleteCassetteCleaning(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteCleaningDelete = new OleDbCommand("DELETE FROM tblCassetteCleanings WHERE cassette_Cleaning_ID = @cassette_Cleaning_ID", dbConnection);
                cmdCassetteCleaningDelete.Parameters.Add(new OleDbParameter("@cassette_Cleaning_ID", Convert.ToInt32(this.txtCassette_Cleaning_ID.Text)));
                if (MessageBox.Show("Are you sure you want to delete all of the data on this cassette cleaning?", "Delete Cassette Cleaning Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdCassetteCleaningDelete.ExecuteNonQuery();
                    MessageBox.Show("Delete Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassetteCleanings", "tblCassetteCleanings");
                tblCassetteCleaningsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void Button_Click_btnInsertCassetteServicing(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteServicingInsert = new OleDbCommand("INSERT INTO tblCassetteServicings(cassette_Barcode, date_Needs_Servicing, date_Back_In_Service, photo_Before_Service, photo_After_Service, notes) VALUES("
                + "@cassette_Barcode, "
                + "@date_Needs_Servicing, "
                + "@date_Back_In_Service, "
                + "@photo_Before_Service, "
                + "@photo_After_Service, "
                + "@notes)", dbConnection);
                cmdCassetteServicingInsert.Parameters.Add(new OleDbParameter("@cassette_Barcode", this.txtCassette_Barcode_tblCS.Text));
                cmdCassetteServicingInsert.Parameters.Add(new OleDbParameter("@date_Needs_Servicing", this.dpCassette_Date_Needs_Servicing.SelectedDate));
                cmdCassetteServicingInsert.Parameters.Add(new OleDbParameter("@date_Back_In_Service", this.dpCassette_Date_Back_In_Service.SelectedDate));
                cmdCassetteServicingInsert.Parameters.Add(new OleDbParameter("@photo_Before_Service", this.txtCassette_Photo_Before_Service.Text));
                cmdCassetteServicingInsert.Parameters.Add(new OleDbParameter("@photo_After_Service", this.txtCassette_Photo_After_Service.Text));
                if (string.IsNullOrEmpty(this.txtCassette_Servicing_Notes.Text))
                    cmdCassetteServicingInsert.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdCassetteServicingInsert.Parameters.Add(new OleDbParameter("@notes", this.txtCassette_Servicing_Notes.Text));

                if (MessageBox.Show("Are you sure you want to add a new cassette servicing?", "Add Cassette Servicing Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdCassetteServicingInsert.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassetteServicings", "tblCassetteServicings");
                tblCassetteServicingsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void Button_Click_btnUpdateCassetteServicing(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteServicingUpdate = new OleDbCommand("UPDATE tblCassetteServicings SET "
                + "cassette_Barcode = @cassette_Barcode, "
                + "date_Needs_Servicing = @date_Needs_Servicing, "
                + "date_Back_In_Service = @date_Back_In_Service, "
                + "photo_Before_Service = @photo_Before_Service, "
                + "photo_After_Service = @photo_After_Service, "
                + "notes = @notes, "
                + "WHERE cassette_Servicing_ID = @cassette_Servicing_ID", dbConnection);
                cmdCassetteServicingUpdate.Parameters.Add(new OleDbParameter("@cassette_Barcode", this.txtCassette_Barcode_tblCS.Text));
                cmdCassetteServicingUpdate.Parameters.Add(new OleDbParameter("@date_Needs_Servicing", this.dpCassette_Date_Needs_Servicing.SelectedDate));
                cmdCassetteServicingUpdate.Parameters.Add(new OleDbParameter("@date_Back_In_Service", this.dpCassette_Date_Back_In_Service.SelectedDate));
                cmdCassetteServicingUpdate.Parameters.Add(new OleDbParameter("@photo_Before_Service", this.txtCassette_Photo_Before_Service.Text));
                cmdCassetteServicingUpdate.Parameters.Add(new OleDbParameter("@photo_After_Service", this.txtCassette_Photo_After_Service.Text));
                if (string.IsNullOrEmpty(this.txtCassette_Servicing_Notes.Text))
                    cmdCassetteServicingUpdate.Parameters.Add(new OleDbParameter("@notes", DBNull.Value));
                else
                    cmdCassetteServicingUpdate.Parameters.Add(new OleDbParameter("@notes", this.txtCassette_Servicing_Notes.Text));
                
                cmdCassetteServicingUpdate.Parameters.Add(new OleDbParameter("@cassette_Servicing_ID", Convert.ToInt32(this.txtCassette_Servicing_ID.Text)));

                if (MessageBox.Show("Are you sure you want to save?", "Update Cassette Servicing Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdCassetteServicingUpdate.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassetteServicings", "tblCassetteServicings");
                tblCassetteServicingsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void Button_Click_btnDeleteCassetteServicing(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdCassetteServicingDelete = new OleDbCommand("DELETE FROM tblCassetteServicings WHERE cassette_Servicing_ID = @cassette_Servicing_ID", dbConnection);
                cmdCassetteServicingDelete.Parameters.Add(new OleDbParameter("@cassette_Servicing_ID", Convert.ToInt32(this.txtCassette_Servicing_ID.Text)));
                if (MessageBox.Show("Are you sure you want to delete all of the data on this cassette servicing?", "Delete Cassette Servicing Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdCassetteServicingDelete.ExecuteNonQuery();
                    MessageBox.Show("Delete Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblCassetteServicings", "tblCassetteServicings");
                tblCassetteServicingsDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }
        
        private void Button_Click_btnUpdateUser(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdUserUpdate = new OleDbCommand("UPDATE tblUsers SET "
                + "employee_ID = @employee_ID, "
                + "last_Name = @last_Name, "
                + "first_Name = @first_Name, "
                + "username = @username, "
                + "[email] = @email, "
                + "security_Level = @security_Level, "
                + "login_Attempts = @login_Attempts "
                + "WHERE employee_ID = @employee_ID", dbConnection);
                cmdUserUpdate.Parameters.Add(new OleDbParameter("@employee_ID", Convert.ToInt32(this.txtEmployee_ID.Text)));
                cmdUserUpdate.Parameters.Add(new OleDbParameter("@last_Name", this.txtLast_Name.Text));
                cmdUserUpdate.Parameters.Add(new OleDbParameter("@first_Name", this.txtFirst_Name.Text));
                cmdUserUpdate.Parameters.Add(new OleDbParameter("@username", this.txtUsername.Text));
                cmdUserUpdate.Parameters.Add(new OleDbParameter("@email", this.txtEmail.Text));
                cmdUserUpdate.Parameters.Add(new OleDbParameter("@security_Level", Convert.ToInt32(this.cmbSecurity_Level.SelectionBoxItem.ToString())));
                cmdUserUpdate.Parameters.Add(new OleDbParameter("@login_Attempts", Convert.ToInt32(this.txtLogin_Attempts.Text)));
                if (MessageBox.Show("Are you sure you want to save?", "Update User Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdUserUpdate.ExecuteNonQuery();
                    MessageBox.Show("Save Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblUsers", "tblUsers");
                tblUsersDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }
        
        private void Button_Click_btnDeleteUser(object sender, RoutedEventArgs e)
        {
            dbConnection.Open();
            try
            {
                OleDbCommand cmdUserDelete = new OleDbCommand("DELETE FROM tblUsers WHERE employee_ID = @employee_ID", dbConnection);
                cmdUserDelete.Parameters.Add(new OleDbParameter("@employee_ID", Convert.ToInt32(this.txtEmployee_ID.Text)));
                if (MessageBox.Show("Are you sure you want to delete all of the data on this user?", "Delete User Record", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    cmdUserDelete.ExecuteNonQuery();
                    MessageBox.Show("Delete Successful");
                }
                else
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                getDataSet("SELECT * FROM tblUsers", "tblUsers");
                tblUsersDataGrid.DataContext = dataSet.Tables[0];
                dbConnection.Close();
            }
        }

        private void btnDateQueryCancel_Click(object sender, RoutedEventArgs e)
        {
            dateInputBoxGrid.Visibility = Visibility.Collapsed;
        }

        private void btnTextQueryCancel_Click(object sender, RoutedEventArgs e)
        {
            textInputBoxGrid.Visibility = Visibility.Collapsed;
        }

        private void btnComboQueryCancel_Click(object sender, RoutedEventArgs e)
        {
            comboInputBoxGrid.Visibility = Visibility.Collapsed;
        }
        
    }
}
