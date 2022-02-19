using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;


namespace NFCBrowser
{


    class SQLiteDB
    {
        public string DatabaseName { set; get; }

        private SQLiteConnection con;


        public SQLiteDB()
        {
            // デフォルト値
            this.DatabaseName = "StationCode.db";
        }
        public SQLiteDB(String dbname)
        {
            // デフォルト値
            this.DatabaseName = dbname;
        }

        /// <summary>
        /// 接続のみ　使い終わったら、廃棄すること
        /// </summary>
        /// <returns></returns>
        public void ConnectDB()
        {
            var sqlConnectionSb = new SQLiteConnectionStringBuilder { DataSource = this.DatabaseName };
            con = new SQLiteConnection(sqlConnectionSb.ToString());
            con.Open();
        }
        public String[] selectStationNameToConnectedDB(String area, String company, String station)
        {
            string[] result = new string[2];


            using (var cmd = new SQLiteCommand(con))
            {
                string sqltext = "select CompanyName,StationName from StationCode where ";
                sqltext += " AreaCode = '" + area + "' ";
                sqltext += " and LineCode = '" + company + "' ";
                sqltext += " and StationCode = '" + station + "' ";
                cmd.CommandText = sqltext;
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read() == true)
                    {
                        result[0] = (string)reader["CompanyName"];
                        result[1] = (string)reader["StationName"];
                    }
                }
            }
            return result;
        }
        public void CloseDB()
        {
            con.Close();
        }

        public String[] selectBusNameToConnectedDB(String company, String station)
        {
            string[] result = new string[2];


            using (var cmd = new SQLiteCommand(con))
            {
                string sqltext = "select BusCompanyName,BusStationName from BusCode where ";
                sqltext += " BusLineCode = '" + company + "' ";
                sqltext += " and BusStationCode = '" + station + "' ";
                cmd.CommandText = sqltext;
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read() == true)
                    {
                        result[0] = (string)reader["BusCompanyName"];
                        result[1] = (string)reader["BusStationName"];
                    }
                }
            }
            if(result[0] == null)
            {
                result[0] = company;
                result[1] = station;
            }
            return result;
        }

        //
        // 駅名検索　[路線名][駅名]の配列を返す　→　connectionを使いまわさないので、使わない。
        //
        /*
        public String[] selectStation(String area, String company, String station)
        {
            string[] result = new string[2];
            var sqlConnectionSb = new SQLiteConnectionStringBuilder { DataSource = this.DatabaseName };

            using (var cn = new SQLiteConnection(sqlConnectionSb.ToString()))
            {
                cn.Open();

                using (var cmd = new SQLiteCommand(cn))
                {
                    string sqltext = "select CompanyName,StationName from StationCode where ";
                    sqltext += " AreaCode = '" + area + "' ";
                    sqltext += " and LineCode = '" + company + "' ";
                    sqltext += " and StationCode = '" + station + "' ";
                    cmd.CommandText = sqltext;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read() == true)
                        {
                            result[0] = (string)reader["CompanyName"];
                            result[1] = (string)reader["StationName"];
                        }
                    }
                }
                
            }
            return result;
        }
        */
        //
        // ユーザーデータベースからユーザー名を取得
        //
        public String selectUser(String uid)
        {
            string result = "";
            var sqlConnectionSb = new SQLiteConnectionStringBuilder { DataSource = this.DatabaseName };

            using (var cn = new SQLiteConnection(sqlConnectionSb.ToString()))
            {
                cn.Open();

                using (var cmd = new SQLiteCommand(cn))
                {
                    string sqltext = "select name from users where ";
                    sqltext += " id = '" + uid + "'";
                    
                    cmd.CommandText = sqltext;
                    using (var reader = cmd.ExecuteReader())
                    {
                        
                        while (reader.Read() == true)
                        {
                            result = (string)reader["name"];
                        }
                    }
                }
            }
            return result;
        }
    }
}
