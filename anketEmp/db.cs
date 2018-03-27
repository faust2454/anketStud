using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.Data;
using System.Windows;

namespace anketEmp
{
    class Db
    {
        public MySqlConnection conn;
        public string serverName = "localhost"; // Адрес сервера (для локальной базы пишите "localhost")
        public string log = "root"; // Имя пользователя
        public string dbName = "anket"; //Имя базы данных
        public string port = "3306"; // Порт для подключения
        public string password = ""; // Пароль для подключения

        public Db(string sName, string l, string dbN, string p, string pass)
        {
            serverName = sName; // Адрес сервера
            log = l; // Имя пользователя
            dbName = dbN; //Имя базы данных
            port = p; // Порт для подключения
            password = pass; // Пароль для подключения
        }

        public Db()
        {
        }

        public bool Connect()
        {
            try
            {
                string connStr = "server=" + serverName +
                    ";user=" + log +
                    ";database=" + dbName +
                    ";port=" + port +
                    ";password=" + password + ";";
                conn = new MySqlConnection(connStr);
                conn.Open();
            }
            catch (Exception)
            {
                return true;
            }
            return false;
        }

        public DataTable DbSelect(string sql)
        {
            int er = 1;
            dbSel:
            try
            {
                MySqlCommand sqlCom = new MySqlCommand(sql, conn);
                sqlCom.ExecuteNonQuery();
                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(sqlCom);
                DataTable dt = new DataTable();
                dataAdapter.Fill(dt);
                return dt;
            }
            catch
            {
                if (Connect()) MessageBox.Show("Ошибка получения данных\nВыполняется попытка №" + er++);
                goto dbSel;
            }
        }
        public bool DbInsert(string sql)
        {
            int er = 1;
            dbSel:
            try
            {
                MySqlCommand sqlCom = new MySqlCommand(sql, conn);
                sqlCom.ExecuteNonQuery();
                return true;
            }
            catch
            {
                if (Connect()) MessageBox.Show("Ошибка передачи данных\nВыполняется попытка №" + er++);
                goto dbSel;
            }
        }
    }
}