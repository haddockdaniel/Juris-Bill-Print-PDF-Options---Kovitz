using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using Microsoft.VisualBasic;
using JurisSVR;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public string message = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);


            if (checkRequisitesAreGood())
            {

                string SQL = "update bc set " + getBillCopySQL() +
                            " from billcopy as bc inner join billto as bt on bt.BillToSysNbr = bc.BilCpyBillTo " +
                            " inner join client as c on c.clisysnbr = bt.BillToCliNbr " +
                            " where (" + getWhereClause() + ") ";


                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Finished Update.", 1, 1);

                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                UpdateStatus("Finished Update.", 0, 1);

            }
            else
            {
                MessageBox.Show(message, "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                message = "";
            }
        }




        private string getBillCopySQL()
        {
            string SQL = "";
            if (radioButtonChangeFrom.Checked)
                SQL = " bc.bilcpyprintformat=1, bc.bilcpyexportformat=0 ";
            else if (radioButtonChangeTo.Checked)
                SQL = " bc.bilcpyprintformat=0, bc.bilcpyexportformat=5 ";
            return SQL;
        }


        private bool checkRequisitesAreGood()
        {
            if (string.IsNullOrEmpty(textBoxBranch.Text) && string.IsNullOrEmpty(textBoxClient.Text))
            {
                message = "Either the Client Code or a Branch must be entered";
                return false;
            }
            if (!radioButtonChangeFrom.Checked && !radioButtonChangeTo.Checked)
            {
                message = "Please select a Change option";
                return false;
            }
            return true;
                
        }


        private string getWhereClause()
        {
            string SQL = ""; 
            string Cli = "'" + textBoxClient.Text.Replace(" ", "").Replace(",", "','") + "'";
            string Bra = "";
            string temp = "";
            if (textBoxBranch.Text.ToLower().Contains("slfe"))
            {
                string sql = "select distinct BRANCH from client where BRANCH like 'SLFe%'";
                DataSet dd = _jurisUtility.RecordsetFromSQL(sql);
                if (dd != null && dd.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow rw in dd.Tables[0].Rows)
                        temp = temp + rw[0].ToString() + ",";
                    temp = temp.TrimEnd(',');
                }
            }
            textBoxBranch.Text = textBoxBranch.Text.ToLower().Replace("slfe", temp);
            Bra = "'" + textBoxBranch.Text.Replace(" ", "").Replace(",", "','") + "'";
            Bra = Bra.TrimEnd(',');

            if (!string.IsNullOrEmpty(textBoxBranch.Text) && !string.IsNullOrEmpty(textBoxClient.Text))
                SQL = " bt.BillToCliNbr in (select clisysnbr from client where clicode in (" + Cli + ")) or c.BRANCH in (" + Bra + ") ";
            else if (!string.IsNullOrEmpty(textBoxClient.Text))
                SQL = " bt.BillToCliNbr in (select clisysnbr from client where clicode  in (" + Cli + ")) ";
            else if (!string.IsNullOrEmpty(textBoxBranch.Text))
                SQL = " c.BRANCH in (" + Bra + ") ";
            return SQL;
        }


        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private void LexisNexisLogoPictureBox_Click(object sender, EventArgs e)
        {
            string sql = "";

            sql = "select top 1 dbo.jfn_FormatClientCode(clicode), billtoclinbr from billcopy as bc inner join billto as bt on bt.BillToSysNbr = bc.BilCpyBillTo " +
                 " inner join client as c on c.clisysnbr = bt.BillToCliNbr " +
                 " where(bt.BillToCliNbr in (select clisysnbr from client where arcodes like 'EM00%')) and bc.bilcpyprintformat = 0 and bc.bilcpyexportformat = 5";
           DataSet dd =  _jurisUtility.RecordsetFromSQL(sql);
            if (dd != null && dd.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow rw in dd.Tables[0].Rows)
                {
                    DialogResult gg = MessageBox.Show("We will now change " + rw[0].ToString() + " to NOT email billing (it currently should be)" + "\r\n" + "Have you verified and want to update it?", "Test", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (gg == DialogResult.Yes)
                    {
                        sql = "update bc set bc.bilcpyprintformat=1, bc.bilcpyexportformat=0 " +
                     " from billcopy as bc inner join billto as bt on bt.BillToSysNbr = bc.BilCpyBillTo " +
                     " inner join client as c on c.clisysnbr = bt.BillToCliNbr " +
                     " where ( bt.BillToCliNbr in (select clisysnbr from client where clisysnbr = " + rw[1].ToString() + "))";
                        _jurisUtility.ExecuteNonQueryCommand(0, sql);

                        MessageBox.Show("Done. Now verify it worked", "Test", MessageBoxButtons.OK, MessageBoxIcon.None);
                    }
                }
                sql = "select top 1 dbo.jfn_FormatClientCode(clicode), billtoclinbr from billcopy as bc inner join billto as bt on bt.BillToSysNbr = bc.BilCpyBillTo " +
             " inner join client as c on c.clisysnbr = bt.BillToCliNbr " +
             " where(bt.BillToCliNbr in (select clisysnbr from client where arcodes like 'EM0%' and arcodes not like 'EM00%')) and bc.bilcpyprintformat = 1 and bc.bilcpyexportformat = 0";
                dd.Clear();
                dd = _jurisUtility.RecordsetFromSQL(sql);
                if (dd != null && dd.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow rw in dd.Tables[0].Rows)
                    {
                        DialogResult gg = MessageBox.Show("We will now change " + rw[0].ToString() + " TO email billing (it currently should NOT be)" + "\r\n" + "Have you verified and want to update it?", "Test", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (gg == DialogResult.Yes)
                        {
                            sql = "update bc set  bc.bilcpyprintformat=0, bc.bilcpyexportformat=5  " +
                         " from billcopy as bc inner join billto as bt on bt.BillToSysNbr = bc.BilCpyBillTo " +
                         " inner join client as c on c.clisysnbr = bt.BillToCliNbr " +
                        " where ( bt.BillToCliNbr in (select clisysnbr from client where clisysnbr = " + rw[1].ToString() + "))";
                            _jurisUtility.ExecuteNonQueryCommand(0, sql);

                            MessageBox.Show("Done. Now verify it worked", "Test", MessageBoxButtons.OK, MessageBoxIcon.None);
                        }
                    }






                }
                else
                    MessageBox.Show("There are no clients with EM0(not 0) that are set FROM email billing", "Error", MessageBoxButtons.OK, MessageBoxIcon.None);





            }
            else
                MessageBox.Show("There are no clients with EM00 that are set TO email billing", "Error", MessageBoxButtons.OK, MessageBoxIcon.None);

        }

        private void buttonBulkSetAllMatters_Click(object sender, EventArgs e)
        {
            List<Client> cliListTo = new List<Client>();
            List<Client> cliListFrom = new List<Client>();
            int total = 0;
            Client cli;
            string sql = "select clisysnbr, ARCodes from client where ARCodes like 'EM%'";
            DataSet dd = _jurisUtility.RecordsetFromSQL(sql);
            if (dd != null && dd.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow rw in dd.Tables[0].Rows)
                {
                    cli = new Client();
                    cli.ID = Convert.ToInt32(rw[0].ToString());
                    cli.code = rw[1].ToString();
                    total++;
                    if (cli.code.StartsWith("EM00"))
                        cliListFrom.Add(cli);
                    else
                        cliListTo.Add(cli);
                }

            }
            string SQL = "";
            int runningTotal = 0;
            foreach (Client cc in cliListFrom)
            {
                runningTotal++;
               SQL = "update bc set bc.bilcpyprintformat=1, bc.bilcpyexportformat=0 " +
                " from billcopy as bc inner join billto as bt on bt.BillToSysNbr = bc.BilCpyBillTo " +
                " inner join client as c on c.clisysnbr = bt.BillToCliNbr " +
                " where ( bt.BillToCliNbr in (select clisysnbr from client where clisysnbr = " + cc.ID.ToString() + "))";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating....", runningTotal, total);
            }
            foreach (Client cc in cliListTo)
            {

                SQL = "update bc set  bc.bilcpyprintformat=0, bc.bilcpyexportformat=5  " +
                 " from billcopy as bc inner join billto as bt on bt.BillToSysNbr = bc.BilCpyBillTo " +
                 " inner join client as c on c.clisysnbr = bt.BillToCliNbr " +
                " where ( bt.BillToCliNbr in (select clisysnbr from client where clisysnbr = " + cc.ID.ToString() + "))";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                UpdateStatus("Updating....", runningTotal, total);
            }

            UpdateStatus("Finished Update.", total, total);
                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                UpdateStatus("Ready to Run...", 0, 1);
            cliListFrom.Clear();
            cliListTo.Clear();

        }

        private void labelDescription_Click(object sender, EventArgs e)
        {

        }

    }
}
