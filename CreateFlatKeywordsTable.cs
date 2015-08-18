using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ProductRepository_Migration
{
    public partial class CreateFlatKeywordsTable : Form
    {
        public CreateFlatKeywordsTable()
        {
            InitializeComponent();
        }

        private void btnFlat_Click(object sender, EventArgs e)
        {
            string query = "SELECT DISTINCT [Document_Id] FROM [Staging].[dbo].[Project_Document_Keyword] ORDER BY 1";
            DataSet universe = getQuery(query);

            DataTable universeDataTable = new DataTable();
            universeDataTable = universe.Tables[0];

            foreach (DataRow galaxyRow in universeDataTable.Rows)
            {
                string subQuery = "SELECT DISTINCT [Keyword_Name] FROM [Staging].[dbo].[Project_Document_Keyword] WHERE [Document_Id] = '" + galaxyRow["Document_Id"].ToString().Replace("'", "") + "' ORDER BY 1";
                DataSet starSystem = getQuery(subQuery);

                DataTable starSystemDataTable = new DataTable();
                starSystemDataTable = starSystem.Tables[0];

                string insertQuery = "INSERT INTO [Staging].[dbo].[Project_Document_Keywords_Transposed] VALUES (" + galaxyRow["Document_Id"].ToString() + ", ";

                foreach (DataRow starSystemRow in starSystemDataTable.Rows)
                {
                    if (starSystemDataTable.Rows.Count == 12) {
                        insertQuery += "'" + starSystemRow["Keyword_Name"].ToString().Replace("'", "") + "')";
                    }else {
                        insertQuery += "'" + starSystemRow["Keyword_Name"].ToString().Replace("'", "") + "', ";
                    }
                }

                if (starSystemDataTable.Rows.Count < 12)
                {

                    int nullsToAdd = 12 - starSystemDataTable.Rows.Count;

                    for (int i = 1; i <= nullsToAdd; i++)
                    {
                        if (i == nullsToAdd)
                        {
                            insertQuery += "'')";
                        }
                        else
                        {
                            insertQuery += "'', ";
                        }
                        
                    }

                }

                executeSQLCommand(insertQuery);

            }
        }

        public static string StringConnection()
        {
            return "Server=Vdbw001tv\\msapps;Database=Staging;Trusted_Connection=True;";
        }

        public static DataSet getQuery(string query)
        {
            using (SqlConnection dbConnection = new SqlConnection(StringConnection()))
            {

                dbConnection.Open();

                SqlDataAdapter objCmd = new SqlDataAdapter(query, dbConnection);
                DataSet objDS = new DataSet();
                objCmd.Fill(objDS, "Data");

                dbConnection.Close();

                return objDS;

            }

        }

        public bool executeSQLCommand(string query)
        {

            using (SqlConnection dbConnection = new SqlConnection(StringConnection()))
            {
                SqlCommand objCmd = default(SqlCommand);

                try
                {
                    objCmd = new SqlCommand(query, dbConnection);

                    dbConnection.Open();
                    objCmd.CommandText = query;
                    objCmd.Connection = dbConnection;
                    objCmd.ExecuteNonQuery();

                    return true;

                }
                catch (Exception ex)
                {
                    return false;
                }
                finally
                {
                    objCmd.Dispose();
                }
            }

        }


    }
}
