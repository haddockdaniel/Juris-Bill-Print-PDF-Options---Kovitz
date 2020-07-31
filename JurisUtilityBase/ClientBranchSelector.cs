using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class ClientBranchSelector : Form
    {
        public ClientBranchSelector(JurisUtility _jurisUtility, string windowText, string tableName)
        {
            InitializeComponent();
            this.Text = windowText;
            tName = tableName;
            dataGridView1.ColumnCount = 1;
            dataGridView2.ColumnCount = 1;
            dataGridView1.Columns[0].Name = "Client Codes ";
            dataGridView2.Columns[0].Name = "Branches ";
            JU = _jurisUtility;
            DataSet client = _jurisUtility.RecordsetFromSQL("select client from " + tableName);
            DataSet branch = _jurisUtility.RecordsetFromSQL("select branch from " + tableName);
            fillRecordSet(client, dataGridView1);
            fillRecordSet(branch, dataGridView2);
        }

        public List<string> clients = new List<string>();
        public List<string> branches = new List<string>();
        JurisUtility JU;
        string tName = "";
        public bool executeChanges = false;


        private void fillRecordSet(DataSet ds, DataGridView dv)
        {
            if (ds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow row in ds.Tables[0].Rows)
                    dv.Rows.Add(row[0].ToString());
            }
            dv.Sort(dv.Columns[0], ListSortDirection.Ascending);

        }




        private void buttonAddRemove_Click(object sender, EventArgs e)
        {
            List<int> selectedIndicies1 = new List<int>();

            foreach (DataGridViewCell c in dataGridView1.SelectedCells)
            {
                c.Value = "";
                selectedIndicies1.Add(c.RowIndex);
            }

            selectedIndicies1 = selectedIndicies1.OrderByDescending(i => i).ToList();

            foreach (int row in selectedIndicies1)
                dataGridView1.Rows.RemoveAt(row);
        }



        private void button1_Click(object sender, EventArgs e)
        {
            List<int> selectedIndicies2 = new List<int>();
            foreach (DataGridViewCell c in dataGridView2.SelectedCells)
            {
                c.Value = "";
                selectedIndicies2.Add(c.RowIndex);
            }
            selectedIndicies2 = selectedIndicies2.OrderByDescending(i => i).ToList();
            foreach (int row in selectedIndicies2)
                dataGridView2.Rows.RemoveAt(row);


        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            string SQL = "truncate table dbo." + tName;
            JU.ExecuteNonQueryCommand(0, SQL);
            for (int i = dataGridView1.Rows.Count - 1; i > -1; --i)
            {
                DataGridViewRow row = dataGridView1.Rows[i];
                if (!row.IsNewRow && !string.IsNullOrEmpty(row.Cells[0].Value.ToString()))
                {
                    clients.Add(row.Cells[0].Value.ToString());
                }
            }

            for (int i = dataGridView2.Rows.Count - 1; i > -1; --i)
            {
                DataGridViewRow row = dataGridView2.Rows[i];
                if (!row.IsNewRow && !string.IsNullOrEmpty(row.Cells[0].Value.ToString()))
                {
                    branches.Add(row.Cells[0].Value.ToString());
                }
            }

                if (clients.Count < branches.Count)
                {
                    int diference = branches.Count - clients.Count;
                    for (int i = 0; i < diference + 1; i++)
                        clients.Add("");
                    for (int a = 0; a < branches.Count; a++)
                    {
                        if (String.IsNullOrEmpty(clients[a]))
                        {
                            SQL = "Insert into " + tName + " (Client, Branch) values ('', '" + branches[a] + "')";
                            JU.ExecuteNonQueryCommand(0, SQL);
                        }
                        else
                        {
                            SQL = "Insert into " + tName + " (Client, Branch) values ('" + clients[a] + "', '" + branches[a] + "')";
                            JU.ExecuteNonQueryCommand(0, SQL);
                        }
                    }
                }
                else if (clients.Count > branches.Count)
                {
                    int diference = clients.Count - branches.Count;
                    for (int i = 0; i < diference + 1; i++)
                        branches.Add("");
                    for (int a = 0; a < clients.Count; a++)
                    {
                        if (String.IsNullOrEmpty(branches[a]))
                        {
                            SQL = "Insert into " + tName + " (Client, Branch) values ('" + clients[a] + "', '')";
                            JU.ExecuteNonQueryCommand(0, SQL);
                        }
                        else
                        {
                            SQL = "Insert into " + tName + " (Client, Branch) values ('" + clients[a] + "', '" + branches[a] + "')";
                            JU.ExecuteNonQueryCommand(0, SQL);
                        }
                    }
                }
                else
                {
                    for (int a = 0; a < clients.Count; a++)
                    {
                        SQL = "Insert into " + tName + " (Client, Branch) values ('" + clients[a] + "', '" + branches[a] + "')";
                        JU.ExecuteNonQueryCommand(0, SQL);
                    }
                }


                executeChanges = false;
            
            this.Hide();


        }

        private void buttonUpdateMatters_Click(object sender, EventArgs e)
        {
            string SQL = "truncate table dbo." + tName;
            JU.ExecuteNonQueryCommand(0, SQL);
            for (int i = dataGridView1.Rows.Count - 1; i > -1; --i)
            {
                DataGridViewRow row = dataGridView1.Rows[i];
                if (!row.IsNewRow && !string.IsNullOrEmpty(row.Cells[0].Value.ToString()))
                {
                    clients.Add(row.Cells[0].Value.ToString());
                }
            }

            for (int i = dataGridView2.Rows.Count - 1; i > -1; --i)
            {
                DataGridViewRow row = dataGridView2.Rows[i];
                if (!row.IsNewRow && !string.IsNullOrEmpty(row.Cells[0].Value.ToString()))
                {
                    branches.Add(row.Cells[0].Value.ToString());
                }
            }

            if (clients.Count < branches.Count)
            {
                int diference = branches.Count - clients.Count;
                for (int i = 0; i < diference + 1; i++)
                    clients.Add("");
                for (int a = 0; a < branches.Count; a++)
                {
                    if (String.IsNullOrEmpty(clients[a]))
                    {
                        SQL = "Insert into " + tName + " (Client, Branch) values ('', '" + branches[a] + "')";
                        JU.ExecuteNonQueryCommand(0, SQL);
                    }
                    else
                    {
                        SQL = "Insert into " + tName + " (Client, Branch) values ('" + clients[a] + "', '" + branches[a] + "')";
                        JU.ExecuteNonQueryCommand(0, SQL);
                    }
                }
            }
            else if (clients.Count > branches.Count)
            {
                int diference = clients.Count - branches.Count;
                for (int i = 0; i < diference + 1; i++)
                    branches.Add("");
                for (int a = 0; a < clients.Count; a++)
                {
                    if (String.IsNullOrEmpty(branches[a]))
                    {
                        SQL = "Insert into " + tName + " (Client, Branch) values ('" + clients[a] + "', '')";
                        JU.ExecuteNonQueryCommand(0, SQL);
                    }
                    else
                    {
                        SQL = "Insert into " + tName + " (Client, Branch) values ('" + clients[a] + "', '" + branches[a] + "')";
                        JU.ExecuteNonQueryCommand(0, SQL);
                    }
                }
            }
            else
            {
                for (int a = 0; a < clients.Count; a++)
                {
                    SQL = "Insert into " + tName + " (Client, Branch) values ('" + clients[a] + "', '" + branches[a] + "')";
                    JU.ExecuteNonQueryCommand(0, SQL);
                }
            }
            executeChanges = true;
            this.Hide();
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            executeChanges = false;
            this.Hide();
        }




    }
}
