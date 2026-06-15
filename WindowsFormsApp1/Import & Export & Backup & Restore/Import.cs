using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Import : Form
    {
        string connectionString;
        string csvPath;

        public Import()
        {
            InitializeComponent();
        }

        private void Import_Load(object sender, EventArgs e)
        {
            LoadTables();
            comboBox1.SelectedIndex = -1;
        }

        void LoadTables()
        {
            try
            {
                Connect connect = new Connect();
                connectionString = connect.ConnectDB();
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    MySqlCommand cmd = new MySqlCommand("SHOW tables;", con);
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    DataTable dt2 = new DataTable();
                    da.Fill(dt2);
                    comboBox1.DataSource = dt2;
                    comboBox1.DisplayMember = dt2.Columns[0].ColumnName;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Ошибка подключения:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Выберите CSV файл";
                ofd.Filter = "CSV файл (*.csv)|*.csv";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    csvPath = ofd.FileName;

                    if (!File.Exists(csvPath))
                    {
                        MessageBox.Show("Файл не найден!", "Ошибка",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    label2.Text = "*файл выбран";
                    label3.Text = csvPath;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите таблицу!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(csvPath) || !File.Exists(csvPath))
            {
                MessageBox.Show("Выберите CSV файл!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string tableName = comboBox1.Text;

            try
            {
                string[] lines = File.ReadAllLines(csvPath, Encoding.UTF8);

                if (lines.Length < 2)
                {
                    MessageBox.Show("CSV пустой!");
                    return;
                }

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    bool tableIsEmpty;

                    using (MySqlCommand checkCmd = new MySqlCommand(
                        $"SELECT EXISTS(SELECT 1 FROM `{tableName}` LIMIT 1);", con))
                    {
                        tableIsEmpty = Convert.ToInt32(checkCmd.ExecuteScalar()) == 0;
                    }

                    string autoIncrementColumn = null;

                    string aiQuery = @"
                            SELECT COLUMN_NAME
                            FROM INFORMATION_SCHEMA.COLUMNS
                            WHERE TABLE_SCHEMA = DATABASE()
                              AND TABLE_NAME = @table
                              AND EXTRA LIKE '%auto_increment%'";

                    using (MySqlCommand aiCmd = new MySqlCommand(aiQuery, con))
                    {
                        aiCmd.Parameters.AddWithValue("@table", tableName);

                        object result = aiCmd.ExecuteScalar();

                        if (result != null)
                            autoIncrementColumn = result.ToString();
                    }

                    string[] headers = lines[0].Split(';');

                    // Ищем колонку Password
                    int passwordColumnIndex = Array.FindIndex(
                        headers,
                        h => h.Trim().Equals("Password",
                            StringComparison.OrdinalIgnoreCase));

                    List<int> validIndexes = new List<int>();
                    List<string> dbColumns = new List<string>();

                    for (int i = 0; i < headers.Length; i++)
                    {
                        string columnName = headers[i].Trim();

                        if (!tableIsEmpty &&
                            !string.IsNullOrEmpty(autoIncrementColumn) &&
                            columnName.Equals(autoIncrementColumn,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        validIndexes.Add(i);
                        dbColumns.Add($"`{columnName}`");
                    }

                    if (dbColumns.Count == 0)
                    {
                        MessageBox.Show("Не найдены колонки для импорта.");
                        return;
                    }
                    int dateColumnIndex = Array.FindIndex(
                                            headers,
                                            h => h.Trim().Equals("Date",
                                                StringComparison.OrdinalIgnoreCase));

                    string columns = string.Join(",", dbColumns);

                    string parameters = string.Join(",",
                        Enumerable.Range(0, validIndexes.Count)
                        .Select(i => $"@p{i}"));

                    string insertQuery =
                        $"INSERT INTO `{tableName}` ({columns}) VALUES ({parameters})";

                    int success = 0;
                    List<string> errors = new List<string>();

                    for (int row = 1; row < lines.Length; row++)
                    {
                        if (string.IsNullOrWhiteSpace(lines[row]))
                            continue;

                        string[] values = lines[row].Split(';');

                        using (MySqlCommand cmd = new MySqlCommand(insertQuery, con))
                        {
                            cmd.CommandTimeout = 60;

                            for (int j = 0; j < validIndexes.Count; j++)
                            {
                                string value = validIndexes[j] < values.Length
                                    ? values[validIndexes[j]].Trim()
                                    : null;

                                if (!string.IsNullOrWhiteSpace(value))
                                {

                                    if (validIndexes[j] == dateColumnIndex)
                                    {
                                        if (DateTime.TryParse(value, out DateTime dt))
                                        {
                                            cmd.Parameters.Add($"@p{j}", MySqlDbType.Date)
                                                .Value = dt.Date;
                                            continue;
                                        }
                                    }

                                    if (headers[validIndexes[j]].Trim().Equals("Sum", StringComparison.OrdinalIgnoreCase) ||
                                                    headers[validIndexes[j]].Trim().Equals("Discount", StringComparison.OrdinalIgnoreCase) ||
                                                    headers[validIndexes[j]].Trim().Equals("TotalSum", StringComparison.OrdinalIgnoreCase))
                                    {
                                        decimal dec = decimal.Parse(
                                            value.Replace(',', '.'),
                                            System.Globalization.CultureInfo.InvariantCulture);

                                        cmd.Parameters.Add($"@p{j}", MySqlDbType.Decimal).Value = dec;
                                        continue;
                                    }

                                    // Хэшируем пароль если нужно
                                    if (validIndexes[j] == passwordColumnIndex)
                                    {
                                        bool isHash =
                                            value.Length == 64 &&
                                            System.Text.RegularExpressions.Regex.IsMatch(
                                                value,
                                                @"^[a-fA-F0-9]{64}$");

                                        if (!isHash)
                                        {
                                            using (var sha = System.Security.Cryptography.SHA256.Create())
                                            {
                                                byte[] hashBytes = sha.ComputeHash(
                                                    Encoding.UTF8.GetBytes(value));

                                                value = BitConverter.ToString(hashBytes)
                                                    .Replace("-", "")
                                                    .ToLower();
                                            }
                                        }
                                        else
                                        {
                                            value = value.ToLower();
                                        }
                                    }

                                    cmd.Parameters.AddWithValue($"@p{j}", value);
                                }
                                else
                                {
                                    cmd.Parameters.AddWithValue($"@p{j}", DBNull.Value);
                                }
                            }

                            try
                            {
                                cmd.ExecuteNonQuery();
                                success++;
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"Строка {row + 1}: {ex.Message}");
                            }
                        }
                    }

                    string message =
                        $"Импорт завершён.\n\n" +
                        $"Успешно: {success}\n" +
                        $"Ошибок: {errors.Count}";

                    if (errors.Count > 0)
                    {
                        message += "\n\nПервые ошибки:\n\n" +
                                   string.Join("\n", errors.Take(10));
                    }

                    MessageBox.Show(
                        message,
                        errors.Count == 0 ? "Успех" : "Завершено с ошибками",
                        MessageBoxButtons.OK,
                        errors.Count == 0
                            ? MessageBoxIcon.Information
                            : MessageBoxIcon.Warning);
                }

                comboBox1.SelectedIndex = -1;
                csvPath = null;
                label3.Text = "";
                label2.Text = "*файл не выбран";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка импорта:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
