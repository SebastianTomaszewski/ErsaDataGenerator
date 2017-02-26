using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MetroFramework.Forms;
 

namespace ErsaDataGenerator
{
    public partial class Form1 : MetroForm
    {

        private static string BaseDirPath => AppDomain.CurrentDomain.BaseDirectory;
        private static string ConfigFilePath => Path.Combine(BaseDirPath, "test.xlsx");

        public Form1()
        {
            InitializeComponent();
            metroDateTime1.Value = metroDateTime1.Value.AddDays(-1);
            metroDateTime4.Value = metroDateTime4.Value.AddDays(-1);

            DB = "CofA_prod";
            SqlIp = @"192.168.0.187\DevIT";
            SqlUser = "sa";
            SqlPass = "b52mkmn5";

            SqlText = "SELECT id,name  FROM test";
        }

        public string DateStart { get; set; }
        public string DateEnd { get; set; }

        public string ConnStr => GenerateConnectionString(SqlUser, SqlPass, SqlIp, DB);
       

        private void metroButton1_Click(object sender, EventArgs e)
        {

            DateStart = metroDateTime1.Value.ToDateString();
            DateEnd = metroDateTime2.Value.ToDateString();
           

           // MessageBox.Show($@"Od {DateStart} do {DateEnd}");

            ClsGeneral dupa = new ClsGeneral();


   //         MessageBox.Show(get_one_text_value());

           dupa.GenerateExcel(ConfigFilePath, GetData());




        }

        #region private variables

        private string sqltext;
        private SqlConnection sqlconn = new SqlConnection();
        private SqlCommand command = new SqlCommand();
        private SqlDataAdapter da = new SqlDataAdapter();
        private string get_one_text_value_string;
        private string db;
        private string sqluser;
        private string sqlpass;
        private string sqlip;
        //private string connectionstring = "Data Source=" + sqlip + ";Persist Security Info=True;User ID=" + sqluser + ";Password=" + sqlpass + ";Database=" + db;
        #endregion

        #region propertis
        public string SqlText { get { return sqltext; } set { sqltext = value; } }

        public string DB { get { return db; } set { db = value; } }
        public string SqlUser { get { return sqluser; } set { sqluser = value; } }
        public string SqlPass { get { return sqlpass; } set { sqlpass = value; } }
        public string SqlIp { get { return sqlip; } set { sqlip = value; } }



        #endregion

        public DataTable GetDt()
        {

           
           


                try
                {
                    if (sqlconn.State != ConnectionState.Closed) sqlconn.Close();
                    sqlconn.ConnectionString = "Data Source=" + sqlip + ";Persist Security Info=True;User ID=" + sqluser + ";Password=" + sqlpass + ";Database=" + db;
                    command.CommandText = sqltext;
                    command.Connection = sqlconn;
                    da.SelectCommand = command;
                    var getDt = new DataTable();
                    sqlconn.Open();
                    da.Fill(getDt);
                    sqlconn.Close();
                    return getDt;
                }
                catch (Exception ex)
                {
                    sqlconn.Close();
                    throw new Exception("blad wykonania GetDt " + command.CommandText + " error: " + ex.Message);
                }
            }

        public DataTable GetData()
        {
            SqlConnection conn = new SqlConnection(ConnStr);
            conn.Open();
            string query = SqlText;
            SqlCommand cmd = new SqlCommand(query, conn);

            DataTable dt = new DataTable("test");
            dt.Load(cmd.ExecuteReader());
            return dt;
        }

        private string GenerateConnectionString(string user, string password, string host, string catalog)
        {
            var sqlBuilder = new SqlConnectionStringBuilder
            {
                Password = password,
                UserID = user,
                DataSource = host,
                InitialCatalog = catalog
            };

            return sqlBuilder.ToString();
        }

        public DataTable PullData()
        {
            var dataTable = new DataTable();
            string connString = ConnStr;
            string query = SqlText;

            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();

            // create data adapter
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable
            da.Fill(dataTable);
            conn.Close();
            da.Dispose();

            return dataTable;
        }


        public string get_one_text_value()
        {
            
                try
                {
                    if (sqlconn.State != ConnectionState.Closed) sqlconn.Close();
                    sqlconn.ConnectionString = "Data Source=" + sqlip + ";Persist Security Info=True;User ID=" + sqluser + ";Password=" + sqlpass + ";Database=" + db;
                    command.CommandText = sqltext;
                    command.Connection = sqlconn;
                    da.SelectCommand = command;
                    DataTable GetDt = new DataTable("Test");
                    sqlconn.Open();
                    da.Fill(GetDt);
                    sqlconn.Close();
                    get_one_text_value_string = GetDt.Rows[0][0].ToString();
                    return get_one_text_value_string;
                }
                catch (Exception ex)
                {
                    sqlconn.Close();
                    throw new Exception("blad wykonania get_one_text_value " + command.CommandText + " error: " + ex.Message);
                }
            }
          


    }
    }

   
