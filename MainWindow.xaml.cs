using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace T4CCWSnapshot
{
    static class Constants
    {
        // database login information is hard-coded below. This is not ideal for security, and
        // it means building a new copy of the app for every client; however, the approach below
        // was selected because factors unique to the client for which this was built created a 
        // need to have the simplest, lowest-friction solution to extracting data from their database

        // database connection information
        public const String dbHost = ""; // i.e. localhost\\instancename";
        public const String dbPort = ""; 
        public const String dbName = "";
        public const String username = "";
        public const String password = "";

        // All systems
        public const String zipFilenameDefault = "CWSnapshot.zip";
        public const String csvPeopleFilename = "individuals.csv";
        public const String csvContributionsFilename = "giving.csv";
    }
    public class Result
    {
        public Boolean Error { get; set; }
        public String ErrorMsg { get; set; }
        public Result()
        {
            this.Error = false;
            this.ErrorMsg = "";
        }
    }

    public class ThreadParameters
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public String SnapshotLocation { get; set; }
        public ThreadParameters(DateTime sd, DateTime ed, String snapshotLocation/*, String pfn, String cfn*/)
        {
            this.StartDate = sd;
            this.EndDate = ed;
            this.SnapshotLocation = snapshotLocation;
        }
    }


    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private String dbHost;
        private String dbPort;
        private String dbName;
        private String username;
        private String password;
        private String PeopleQueryString;
        private String PeopleFilename;
        private String ContributionsQueryString;
        private String ContributionsFilename;
        private String LogFilename;

        public MainWindow()
        {
            InitializeComponent();

            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            InitializeDefaults();

            if (File.Exists(this.LogFilename))
            {
                File.Delete(this.LogFilename);
                LogInfo("Deleted existing logfile: " + this.LogFilename);
            }
        }

        private void About_Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("T4CCWSnapshot's icon incorporates work from Good Ware at www.flaticon.com", "About");
        }

        private void Change_Snapshot_Location_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = System.IO.Path.GetFileName(this.snapshotLocation.Text); // Default file name
            dlg.DefaultExt = ".zip"; // Default file extension
            dlg.Filter = "Zip file (.zip)|*.zip"; // Filter files by extension
            dlg.InitialDirectory = System.IO.Path.GetDirectoryName(this.snapshotLocation.Text);
            dlg.OverwritePrompt = true;

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                this.snapshotLocation.Text = dlg.FileName;
            }
        }

        private void Create_Snapshot_Click(object sender, RoutedEventArgs e)
        {
            
            Thread thread = new Thread(DoBackgroundQueryAndFileWrite);

            var parms = new ThreadParameters((System.DateTime)this.startDate.SelectedDate, (System.DateTime)this.endDate.SelectedDate, this.snapshotLocation.Text/*, this.PeopleFilename, this.ContributionsFilename*/);

            thread.Start(parms);
        }

        private void DoBackgroundQueryAndFileWrite(object data)
        {
            ThreadParameters parms = (ThreadParameters)data;

            DateTime startDate = parms.StartDate;
            DateTime endDate = parms.EndDate;
            String snapshotLocation = parms.SnapshotLocation;
            Result result = new Result();

            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                this.overlay.Visibility = System.Windows.Visibility.Visible;
                this.statusLabel.Content = "Querying People";
            }));
            
            result = WritePeopleToFile(startDate, endDate);

            if (!result.Error)
            {
                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    this.statusLabel.Content = "Querying Contributions";
                }));
                result = WriteContributionsToFile(startDate, endDate);
            }

            if (!result.Error)
            {
                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    this.statusLabel.Content = "Zipping Files";
                }));
                result = ZipFiles(snapshotLocation);
            }

            if (result.Error)
            {
                LogError(result.ErrorMsg);
                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    this.statusLabel.Content = result.ErrorMsg;
                }));
            }
            else
            {
                LogInfo("Snapshot created successfully");
                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    this.statusLabel.Content = "Snapshot Created.";
                    MessageBox.Show("Snapshot file (" + snapshotLocation + ") ready for upload to Tools4Church.");
                }));
            }

            if (File.Exists(this.PeopleFilename))
            {
                LogInfo("Deleting temporary people file: " + this.PeopleFilename);
                File.Delete(this.PeopleFilename);
            }
            if (File.Exists(this.ContributionsFilename))
            {
                LogInfo("Deleting temporary contributions file: " + this.ContributionsFilename);
                File.Delete(this.ContributionsFilename);
            }

            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                this.overlay.Visibility = System.Windows.Visibility.Hidden;
            })); 
        }

        private SqlConnection GetDBConnection()
        {
            SqlConnection c = new SqlConnection(@"Data Source=" + this.dbHost + "," + this.dbPort + ";Initial Catalog=" + this.dbName + ";User ID=" + this.username + ";Password=" + this.password);

            return c;
        }

        private void InitializeDefaults()
        {
            this.dbHost = Constants.dbHost;
            this.dbPort = Constants.dbPort;
            this.dbName = Constants.dbName;
            this.username = Constants.username;
            this.password = Constants.password;

            this.endDate.SelectedDate = DateTime.Now;
            this.startDate.SelectedDate = (DateTime.Now).Subtract(new TimeSpan(90, 0, 0, 0));
            this.snapshotLocation.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + Constants.zipFilenameDefault;

            this.LogFilename = System.IO.Path.GetTempPath() + "T4CCWSnapshot.log";

            this.ContributionsQueryString = @"
                SELECT 
                    CONVERT(varchar, m.DateOccurred, 23) as contributionDate, 
		            cwig.FamilyID as cwFamilyID,
		            m.giverID as cwGiverID,
		            a.AccountID as cwFundID,
		            a.AccountName as fundDescr, 
		            m.PaymentMethod + 1 as givingMethodID,
		            case m.PaymentMethod 
                        When 0 then 'None' 
                        When 1 then 'Computer Check' 
                        When 2 then 'Manual Check'
                        when 3 then 'Check' 
                        When 4 then 'EFT' 
                        when 5 then 'Cash' 
                        when 6 then 'Debit Card' 
                        when 7 then 'Credit Card'
                        when 8 then 'Stock' 
                        When 9 then 'In Kind' 
                        else 'Unknown'
                    end as givingMethodDescr,
		            CASE WHEN d.PledgeID IS NULL THEN '' ELSE convert(nvarchar(36), d.PledgeID) END as cwPledgeDriveID,
		            CASE WHEN d.PledgeID IS NULL THEN '' ELSE a.AccountName + ' - ' + (convert(varchar(20), p.BeginDate, 101) + ' - ' + convert(varchar(20), p.EndDate, 101)) END as pledgeDriveDescr,
                SUM(d.amount) as amount
                from DonationMaster m
                Join DonationDetail d on m.UniqueID = d.DonationMasterID
                Join GivingAccounts a on d.AccountID = a.AccountID
                left OUTER join Pledge p on p.UniqueID = d.PledgeID
                LEFT OUTER JOIN CWIndividualGivers cwig ON cwig.UniqueID = m.GiverID
                WHERE m.dateposted is not null
                  AND d.amount <> 0
                  AND m.Reversed = 0
                  and m.DateOccurred between @StartDate and @EndDate
                  AND NOT EXISTS (
	  	            SELECT 'X'
	  	            FROM DonationMaster m1
	  	            WHERE m1.Corrected = m.TransactionNumber
	              )
                GROUP BY
                    CONVERT(varchar, m.DateOccurred, 23),
		            cwig.FamilyID,
		            m.giverID,
		            a.AccountID,
		            a.AccountName,
		            m.PaymentMethod,
		            case m.PaymentMethod 
                        When 0 then 'None' 
                        When 1 then 'Computer Check' 
                        When 2 then 'Manual Check'
                        when 3 then 'Check' 
                        When 4 then 'EFT' 
                        when 5 then 'Cash' 
                        when 6 then 'Debit Card' 
                        when 7 then 'Credit Card'
                        when 8 then 'Stock' 
                        When 9 then 'In Kind' 
                    end,
		            CASE WHEN d.PledgeID IS NULL THEN '' ELSE convert(nvarchar(36), d.PledgeID) END,
		            CASE WHEN d.PledgeID IS NULL THEN '' ELSE a.AccountName + ' - ' + (convert(varchar(20), p.BeginDate, 101) + ' - ' + convert(varchar(20), p.EndDate, 101)) END
                HAVING SUM(d.amount) <> 0            
            ";

            this.PeopleQueryString = @"
                SELECT  
                    cwig.UniqueID as cwPersonID,
		            cwig.FamilyID as cwFamilyID,
		            ISNULL(cwig.FirstName, '') as firstName,
		            ISNULL(cwig.MiddleName, '') as middleName,
		            ISNULL(cwig.NickName, '') as goesByName,
		            ISNULL(cwig.LastName, '') as lastName,
		            CASE 
			            WHEN cwig.Birthdate IS NULL OR CHARINDEX('?', cwig.Birthdate) > 0 OR LEN(cwig.Birthdate) <> 10 THEN ''
			            ELSE SUBSTRING(cwig.birthdate, 7, 4) + '-' + SUBSTRING(cwig.birthdate, 1, 2) + '-' + SUBSTRING(cwig.birthdate, 4, 2)
		            END as birthdate,
		            CASE 
			            WHEN cwig.Definable1 = '3E5CD732-4A68-4A26-9F23-5C683EE582E2' THEN 'f'
			            WHEN cwig.Definable1 = '257877F9-E1BD-488B-B6E6-B327B8507F4E' THEN 'm'
			            ELSE 'u'
		            END as gender,
		            ISNULL(mc1.description, '') as memberStatus,
		            ISNULL(mc2.description,'') as familyPosition,
		            CASE 
			            WHEN cwig.PermanentEmail IS NULL THEN ''
			            ELSE CASE CHARINDEX('|', cwig.PermanentEmail) WHEN 0 THEN cwig.PermanentEmail ELSE SUBSTRING(cwig.PermanentEmail, 1, CHARINDEX('|', cwig.PermanentEmail)-1) END 
		            END as emailAddress,
		            CASE 
			            WHEN cwig.Definable20 IS NULL THEN ''
			            ELSE CASE CHARINDEX('|', cwig.Definable20) WHEN 0 THEN cwig.Definable20 ELSE SUBSTRING(cwig.Definable20, 1, CHARINDEX('|', cwig.Definable20)-1) END 
		            END as mobilePhone,
		            CASE 
			            WHEN cwfg.HomePhone IS NULL THEN ''
			            ELSE CASE CHARINDEX('|', cwfg.HomePhone) WHEN 0 THEN cwfg.HomePhone ELSE SUBSTRING(cwfg.HomePhone, 1, CHARINDEX('|', cwfg.HomePhone)-1) END 
		            END as homePhone,
		            CASE 
			            WHEN cwig.WorkPhone IS NULL THEN ''
			            ELSE CASE CHARINDEX('|', cwig.WorkPhone) WHEN 0 THEN cwig.WorkPhone ELSE SUBSTRING(cwig.WorkPhone, 1, CHARINDEX('|', cwig.WorkPhone)-1) END 
		            END as workPhone,
		            ISNULL(addr.address1, '') as homeStreet,
		            ISNULL(addr.city, '') as homeCity,
		            ISNULL(addr.state, '') as homeState,
		            ISNULL(addr.zip, '') as homeZip
                FROM CWIndividualGivers cwig 
                JOIN CWFamilyGivers cwfg ON cwfg.UniqueID = cwig.FamilyID
                LEFT OUTER JOIN MembershipCodes mc1 on mc1.UniqueID = cwig.StatusID 
                LEFT OUTER JOIN MembershipCodes mc2 on mc2.UniqueID = cwig.FamilyRelID
                LEFT OUTER JOIN Addresses addr ON addr.UniqueID = (select MIN (addr1.UniqueID) FROM Addresses addr1 WHERE addr1.EntityID = cwfg.UniqueID and addr1.AddressType IN (0, 3)) 
/* uncomment this section to retrieve only the people that made contributions 
                WHERE cwig.FamilyID IN (
  	                SELECT DISTINCT cwig2.FamilyID
  	                FROM CWIndividualGivers cwig2
  	                JOIN DonationMaster dm ON dm.GiverID = cwig2.UniqueID
  	                JOIN DonationDetail dd ON dd.DonationMasterID = dm.UniqueID
	                WHERE dm.dateposted is not null 
	                    AND dd.amount <> 0
	                    AND dm.Reversed = 0
	                    and dm.DateOccurred between @StartDate and @EndDate
	                    AND NOT EXISTS (
	  	                    SELECT 'X'
	  	                    FROM DonationMaster dm1
	  	                    WHERE dm1.Corrected = dm.TransactionNumber
	                    )
                    )
*/
                UNION
                SELECT  g.UniqueID as cwPersonID,
		            NULL as cwFamilyID,
		            '' as firstName,
		            '' as middleName,
		            '' as goesByName,
		            ISNULL(g.GroupName, 'Group #: ' + convert(nvarchar(36), g.UniqueID)) as lastName,
		            NULL as birthdate,
		            'u' as gender,
		            '' as memberStatus,
		            '' as familyPosition,
		            '' as emailAddress,
		            '' as mobilePhone,
		            '' as homePhone,
		            '' as workPhone,
		            '' as homeStreet,
		            '' as homeCity,
		            '' as homeState,
		            '' as homeZip
                FROM DonationMaster m 
	                JOIN DonationDetail d on m.UniqueID = d.DonationMasterID 
	                JOIN Groups g ON g.UniqueID = m.GiverID
                WHERE m.dateposted is not null 
                    AND d.amount <> 0
                    AND m.Reversed = 0
                    and m.DateOccurred between @StartDate and @EndDate
                    AND NOT EXISTS (
  	                    SELECT 'X'
  	                    FROM DonationMaster m1
  	                    WHERE m1.Corrected = m.TransactionNumber
                    )
                UNION
                SELECT  mg.UniqueID as cwPersonID,
		            NULL as cwFamilyID,
		            '' as firstName,
		            '' as middleName,
		            '' as goesByName,
		            ISNULL(mg.Description, 'Membership Group #: ' + convert(nvarchar(36), mg.UniqueID)) as lastName,
		            NULL as birthdate,
		            'u' as gender,
		            '' as memberStatus,
		            '' as familyPosition,
		            '' as emailAddress,
		            '' as mobilePhone,
		            '' as homePhone,
		            '' as workPhone,
		            '' as homeStreet,
		            '' as homeCity,
		            '' as homeState,
		            '' as homeZip
                FROM DonationMaster m 
	                JOIN DonationDetail d on m.UniqueID = d.DonationMasterID 
	                JOIN MembershipGroups mg ON mg.UniqueID = m.GiverID
                WHERE m.dateposted is not null 
                    AND d.amount <> 0
                    AND m.Reversed = 0
                    and m.DateOccurred between @StartDate and @EndDate
                    AND NOT EXISTS (
  	                    SELECT 'X'
  	                    FROM DonationMaster m1
  	                    WHERE m1.Corrected = m.TransactionNumber
                    )
            ";
        }

        private void LogError(String message)
        {
            using (StreamWriter w = File.AppendText(this.LogFilename))
            {
                w.Write("\r\n" + DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " - Error: " + message);
            }
        }

        private void LogInfo(String message)
        {
            using (StreamWriter w = File.AppendText(this.LogFilename))
            {
                w.Write("\r\n" + DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " - Info: " + message);
            }
        }

        private Result WriteContributionsToFile(DateTime startDate, DateTime endDate)
        {
            Result result = new Result();
            int readCount = 0;

            try
            {
                using (SqlConnection connection = GetDBConnection())
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(this.ContributionsQueryString, connection);
                    command.Parameters.Add("@StartDate", SqlDbType.Date).Value = startDate;
                    command.Parameters.Add("@EndDate", SqlDbType.Date).Value = endDate;
                    command.Prepare();
                    SqlDataReader dataReader = command.ExecuteReader();

                    Boolean success = dataReader.Read();

                    if (success)
                    {
                        this.ContributionsFilename = System.IO.Path.GetTempFileName();

                        using (StreamWriter sw = File.CreateText(this.ContributionsFilename))
                        {
                            using (CsvHelper.CsvWriter writer = new CsvHelper.CsvWriter(sw, CultureInfo.InvariantCulture))
                            {
                                while (success)
                                {
                                    readCount++;

                                    int i = 0;
                                    var contributions = new List<T4CContribution> { };
                                    var t4c = new T4CContribution { };
                                    t4c.ContributionDate  = dataReader.GetString(i++);
                                    t4c.CWFamilyID        = !dataReader.IsDBNull(i) ? dataReader.GetGuid(i).ToString() : "";
                                    i++;
                                    t4c.CWGiverID         = dataReader.GetGuid(i++).ToString();
                                    t4c.CWFundID          = dataReader.GetGuid(i++).ToString();
                                    t4c.FundDescr         = dataReader.GetString(i++);
                                    t4c.GivingMethodID    = dataReader.GetInt32(i++).ToString();
                                    t4c.GivingMethodDescr = dataReader.GetString(i++);
                                    t4c.CWPledgeDriveID   = dataReader.GetString(i++);
                                    t4c.PledgeDriveDescr  = dataReader.GetString(i++);
                                    t4c.Amount            = dataReader.GetDecimal(i++).ToString();

                                    contributions.Add(t4c);

                                    writer.WriteRecords(contributions);

                                    success = dataReader.Read();
                                }
                                dataReader.Close();
                            }
                            sw.Close();
                        }
                        LogInfo("Read and wrote " + readCount + " contribution records");
                    }
                    else
                    {
                        result.Error = true;
                        result.ErrorMsg = "Check date range. Unsuccessful read of contribution data from Church Windows.";
                        LogInfo("Success != true on DB read of contribution data.");
                    }
                }
            }
            catch (Exception e)
            {
                result.Error = true;
                result.ErrorMsg = "Error reading from database";
                LogInfo("DB error reading contribution data. Message: " + e.Message);
            }

            if (!result.Error && readCount == 0)
            {
                result.Error = true;
                result.ErrorMsg = "Check date range. No records returned when querying contribution data.";
                LogInfo("readCount == 0 on DB read of contribution data.");
            }

            return result;
        }
        
        private Result WritePeopleToFile(DateTime startDate, DateTime endDate)
        {
            Result result = new Result();
            int readCount = 0;

            try
            {
                using (SqlConnection connection = GetDBConnection())
                {
                    connection.Open();

                    SqlCommand command = new SqlCommand(this.PeopleQueryString, connection);
                    command.Parameters.Add("@StartDate", SqlDbType.Date).Value = startDate;
                    command.Parameters.Add("@EndDate", SqlDbType.Date).Value = endDate;
                    command.Prepare();

//                    SqlCommand command = new SqlCommand("select count(*) as count from cwindividualgivers", connection);
                    SqlDataReader dataReader = command.ExecuteReader();

                    Boolean success = dataReader.Read();

                    if (success)
                    {
                        this.PeopleFilename = System.IO.Path.GetTempFileName();

                        using (StreamWriter sw = File.CreateText(this.PeopleFilename))
                        {
                            using (CsvHelper.CsvWriter writer = new CsvHelper.CsvWriter(sw, CultureInfo.InvariantCulture))
                            {
                                while (success)
                                {
                                    readCount++;
                                    //LogInfo("Count = " + dataReader.GetInt32(0).ToString());
                                    //break;
                                    int i = 0;
                                    var person = new List<T4CPerson> {};
                                    var t4cPerson = new T4CPerson { };
                                    t4cPerson.CWPersonID = dataReader.GetGuid(i++).ToString();
                                    t4cPerson.CWFamilyID = !dataReader.IsDBNull(i) ? dataReader.GetGuid(i).ToString() : "";
                                    i++;
                                    t4cPerson.FirstName         = dataReader.GetString(i++);
                                    t4cPerson.MiddleName        = dataReader.GetString(i++);
                                    t4cPerson.GoesByName        = dataReader.GetString(i++);
                                    t4cPerson.LastName          = dataReader.GetString(i++);
                                    t4cPerson.Birthdate         = !dataReader.IsDBNull(i) ? dataReader.GetString(i) : "";
                                    i++;
                                    t4cPerson.Gender            = dataReader.GetString(i++);
                                    t4cPerson.MemberStatus      = dataReader.GetString(i++);
                                    t4cPerson.FamilyPosition    = dataReader.GetString(i++);
                                    t4cPerson.EmailAddress      = dataReader.GetString(i++);
                                    t4cPerson.MobilePhone       = dataReader.GetString(i++);
                                    t4cPerson.HomePhone         = dataReader.GetString(i++);
                                    t4cPerson.WorkPhone         = dataReader.GetString(i++);
                                    t4cPerson.HomeStreet        = dataReader.GetString(i++);
                                    t4cPerson.HomeCity          = dataReader.GetString(i++);
                                    t4cPerson.HomeState         = dataReader.GetString(i++);
                                    t4cPerson.HomeZip           = dataReader.GetString(i++);

                                    person.Add(t4cPerson);
                                    
                                    writer.WriteRecords(person);

                                    success = dataReader.Read();
                                }
                                dataReader.Close();
                            }
                            sw.Close();
                        }

                        LogInfo("Read and wrote " + readCount + " people records");
                    }
                    else
                    {
                        result.Error = true;
                        result.ErrorMsg = "Unsuccessful read of people data from Church Windows.";
                        LogInfo("Success != true on DB read of people data.");
                    }
                }
            }
            catch (Exception e)
            {
                result.Error = true;
                result.ErrorMsg = "Error reading from database";
                LogInfo("DB error reading people data. Message: " + e.Message);
            }

            if (!result.Error && readCount == 0)
            {
                result.Error = true;
                result.ErrorMsg = "No records returned when querying people data.";
                LogInfo("readCount == 0 on DB read of people data.");
            }

            return result;
        }

        private Result ZipFiles(String snapshotLocation)
        {
            Result result = new Result();

            try
            {
                if (File.Exists(snapshotLocation))
                {
                    LogInfo("Deleting existing zipfile: " + snapshotLocation);
                    File.Delete(snapshotLocation);
                }
                LogInfo("Creating zipfile: " + snapshotLocation);
                using (ZipArchive newFile = ZipFile.Open(snapshotLocation, ZipArchiveMode.Create))
                {
                    LogInfo("Created zipfile: " + snapshotLocation);
                    newFile.CreateEntryFromFile(this.PeopleFilename, Constants.csvPeopleFilename);
                    LogInfo("Added people to zip file from: " + this.PeopleFilename);
                    newFile.CreateEntryFromFile(this.ContributionsFilename, Constants.csvContributionsFilename);
                    LogInfo("Added contributions to zip file from: " + this.ContributionsFilename);
                }
            }
            catch (Exception e)
            {
                result.Error = true;
                result.ErrorMsg = "Error zipping files";
            }
            return result;
        }
    }
}

    
