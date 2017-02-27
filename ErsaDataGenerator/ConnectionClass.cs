using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ErsaDataGenerator.Properties;
using static System.String;

namespace ErsaDataGenerator
{
    public class ConnectionClass
    {
        //private string _connStr;


        public string ConnectionString { get; set; }

        #region ctr

        public ConnectionClass() 
        {
        }
        public ConnectionClass(string conectionString)
        {
            if (IsNullOrEmpty(conectionString)) throw new ArgumentException(Resources.Value_cant_be_null, nameof(conectionString));

            ConnectionString = conectionString;
        }
        public ConnectionClass(string host,string db,string user,string pass)
        {
            if (IsNullOrEmpty(host)) throw new ArgumentException(Resources.Value_cant_be_null, nameof(host));
            if (IsNullOrEmpty(db))   throw new ArgumentException(Resources.Value_cant_be_null, nameof(db));
            if (IsNullOrEmpty(user)) throw new ArgumentException(Resources.Value_cant_be_null, nameof(user));
            if (IsNullOrEmpty(pass)) throw new ArgumentException(Resources.Value_cant_be_null, nameof(pass));

            ConnectionString = GetConnectionString(host, db, user, pass);
        }

        #endregion


        private static string GetConnectionString(string host, string db, string user, string pass)
        {
            var sqlBuilder = new SqlConnectionStringBuilder
            {
                Password = pass,
                UserID = user,
                DataSource = host,
                InitialCatalog = db
            };

            return sqlBuilder.ToString();
        }


        public DataTable GetDataQueryTable(string query)
        {
            if (IsNullOrEmpty(query)) throw new ArgumentException(Resources.Value_cant_be_null, nameof(query));
            if (ConnectionString == null) throw new ArgumentNullException(nameof(ConnectionString));

            using (var conn = new SqlConnection(ConnectionString))
            {
                conn.Open();

                DataTable dt;
                using (var cmd = new SqlCommand(query, conn))
                {
                    dt = new DataTable();
                    dt.Load(cmd.ExecuteReader());
                }
                return dt;
            }
        }








    }
}
