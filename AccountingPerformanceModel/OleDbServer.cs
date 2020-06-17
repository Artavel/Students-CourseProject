using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.OleDb;

namespace AccountingPerformanceModel
{
    // Класс для работы с базой данных OleDB сервера
    public class OleDbServer
    {
        public string Connection { get; set; } = string.Empty; // строка подключения
        public string LastError { get; set; } = string.Empty; // последняя ошибка

        // Выполнить SQL-запрос
        public bool ExecSql(string sql, Dictionary<string, object> columns = null)
        {
            bool result = false;
            using (var con = new OleDbConnection(Connection))
            {
                con.Open();
                using (OleDbCommand command = new OleDbCommand(sql, con))
                {
                    try
                    {
                        if (columns != null)
                        {
                            foreach (var key in columns.Keys)
                                command.Parameters.AddWithValue($"@{key}", columns[key]);
                        }
                        var rows = command.ExecuteNonQuery();
                        LastError = "";
                        result = true;
                    }
                    catch (Exception e)
                    {
                        LastError = e.Message;
                        return false;
                    }
                }
            }
            return result;
        }

        // Проверка на существование записи с ключом
        public bool KeyRecordExists(string table, string keyName, Guid valueValue)
        {
            bool result = false;
            using (var con = new OleDbConnection(Connection))
            {
                con.Open();
                var sql = $"SELECT COUNT(*) FROM `{table}` WHERE `{keyName}` = @{keyName}";
                using (OleDbCommand command = new OleDbCommand(sql, con))
                {
                    command.Parameters.AddWithValue($"@{keyName}", "P"+valueValue.ToString());
                    try
                    {
                        var value = (int)command.ExecuteScalar();
                        LastError = "";
                        result = value > 0;
                    }
                    catch (Exception e)
                    {
                        LastError = e.Message;
                        return false;
                    }
                }
            }
            return result;
        }

        // Запрос на вставку данных
        public bool InsertInto(string table, Dictionary<string, object> columns)
        {
            // формирование запроса для изменения
            var props = new List<string>();
            var values = new List<string>();
            foreach (var key in columns.Keys)
            {
                props.Add($"`{key}`");
                var value = columns[key];
                values.Add($"@{key}");
            }
            var sql = $"INSERT INTO `{table}` ({string.Join(",", props)}) VALUES ({string.Join(",", values)})";
            return ExecSql(sql, columns);
        }

        // Запрос на изменение данных
        public bool UpdateInto(string table, Dictionary<string, object> columns)
        {
            return DeleteInto(table, columns) ? InsertInto(table, columns) : false;
        }

        // Удаление всех записей из таблицы
        public bool DeleteInto(string table)
        {
            // формирование запроса для удаления
            return ExecSql($"DELETE FROM `{table}`");
        }

        // Удаление конкретной записи
        public bool DeleteInto(string table, Dictionary<string, object> columns)
        {
            // формирование запроса для удаления
            var indexName = columns.Keys.First();
            return ExecSql($"DELETE FROM `{table}` WHERE `{indexName}`=@{indexName}", columns);
        }

        // Получение набора данных из таблицы
        public DataSet GetRows(string table)
        {
            using (var con = new OleDbConnection(Connection))
            {
                using (var da = new OleDbDataAdapter($"SELECT * FROM `{table}`", con))
                {
                    var ds = new DataSet();
                    try
                    {
                        da.Fill(ds, table);
                        LastError = "";
                    }
                    catch (Exception ex)
                    {
                        LastError = ex.Message;
                    }
                    return ds;
                }
            }
        }

    }
}
