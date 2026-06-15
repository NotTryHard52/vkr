using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Patient : Form
    {
        string connectionString; // Строка подключения к базе данных
        DataTable patientTable; // Таблица для хранения данных пациентов
        public event Action<int, string> PatientSelected; // Событие для передачи выбранного пациента
        int selectedId = -1; // Id выбранного пациента
        private bool openedFromTalon = false; // Флаг, был ли вызов формы из создания талона
        // Текущий не замаскированный пациент (по id)
        private int currentlyUnmaskedPatientId = -1;
        // Резерв маскированных значений для восстановления
        private System.Collections.Generic.Dictionary<int, (string Name, string Lastname, string Phone, string Policy)> maskedBackup
            = new System.Collections.Generic.Dictionary<int, (string, string, string, string)>();
        private Timer inactivityTimer;
        private DateTime lastActivityTime;
        private const int timeoutSeconds = 120;
        int currentPage = 1;
        int pageSize = 10;
        int totalRecords = 0;
        int totalPages = 1;
        public event Action OnSessionExpired;

        public Patient(bool fromTalon = false)
        {
            try
            {
                InitializeComponent();

                openedFromTalon = fromTalon;

                // Кнопка "Выбрать" отображается только если форма открыта из талона
                button4.Visible = openedFromTalon;
                // При изменении размера таблицы перезагружаем только если изменилось количество видимых строк
                dataGridView1.SizeChanged += DataGridView1_SizeChanged;

                inactivityTimer = new Timer();
                inactivityTimer.Interval = 1000; // проверка каждую секунду
                inactivityTimer.Tick += InactivityTimer_Tick;
                inactivityTimer.Start();

                lastActivityTime = DateTime.Now;

                // отслеживание активности
                RegisterActivityHandlers(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации формы:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResetActivity(object sender, EventArgs e)
        {
            lastActivityTime = DateTime.Now;
        }

        private void InactivityTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if ((DateTime.Now - lastActivityTime).TotalSeconds >= timeoutSeconds)
                {
                    inactivityTimer.Stop();
                    OnSessionExpired?.Invoke();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка таймера:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DataGridView1_SizeChanged(object sender, EventArgs e)
        {
            try
            {
                int newPageSize = CalculatePageSize();
                if (newPageSize != pageSize)
                {
                    // сбрасываем на первую страницу при изменении количества строк на странице
                    currentPage = 1;
                    LoadPatient();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при изменении размера таблицы:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Загрузка формы
        private void Patient_Load(object sender, EventArgs e)
        {
            try
            {
                comboBox2.SelectedIndex = 0; // Установка сортировки по умолчанию
                LoadPatient(); // Загрузка данных пациентов
                var hoverEffect = new HoverDataGridView(dataGridView1); // Визуальный эффект наведения на строки

                // Обработчик двойного клика — централизованный
                dataGridView1.CellDoubleClick += DataGridView1_CellDoubleClick;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке формы:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int CalculatePageSize()
        {
            try
            {
                int rowHeight = dataGridView1.RowTemplate.Height;
                int headerHeight = dataGridView1.ColumnHeadersHeight;

                int availableHeight = dataGridView1.DisplayRectangle.Height;

                int rows = availableHeight / rowHeight;

                return Math.Max(1, rows - 1);
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Ошибка при расчете количества строк на странице:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 10; // Возвращаем значение по умолчанию в случае ошибки
            }
        }

        // Загрузка данных пациентов из базы данных
        private void LoadPatient()
        {
            try
            {
                pageSize = CalculatePageSize();
                Connect connect = new Connect();
                connectionString = connect.ConnectDB();

                string filterSql = BuildFilterSql();
                string sortSql = GetSortSql();

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    // считаем записи
                    totalRecords = GetTotalCount(filterSql);
                    totalPages = (int)Math.Ceiling((double)totalRecords / pageSize);

                    if (currentPage > totalPages)
                        currentPage = totalPages == 0 ? 1 : totalPages;

                    int offset = (currentPage - 1) * pageSize;

                    string query = $@"
                                    SELECT *
                                    FROM Patients
                                    {filterSql}
                                    {sortSql}
                                    LIMIT @offset, @pageSize;
                                ";

                    MySqlCommand cmd = new MySqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@offset", offset);
                    cmd.Parameters.AddWithValue("@pageSize", pageSize);

                    patientTable = new System.Data.DataTable();
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    da.Fill(patientTable);

                    dataGridView1.DataSource = null;
                    // Маскировка
                    MaskNameAndPatronymic();

                    dataGridView1.DataSource = patientTable;

                    // Снимаем выделение строк после привязки чтобы при изменении размера не подсвечивались все строки
                    try
                    {
                        dataGridView1.ClearSelection();
                        // Сброс текущей ячейки, если возможно
                        if (dataGridView1.CurrentCell != null)
                            dataGridView1.CurrentCell = null;
                    }
                    catch
                    {
                        // Игнорируем возможные ошибки при сбросе CurrentCell
                    }

                    // Заголовки
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns[1].HeaderText = "Фамилия";
                    dataGridView1.Columns[2].HeaderText = "Имя";
                    dataGridView1.Columns[3].HeaderText = "Отчество";
                    dataGridView1.Columns[4].HeaderText = "Дата рождения";
                    dataGridView1.Columns[5].HeaderText = "Телефон";
                    dataGridView1.Columns[6].HeaderText = "Полис";

                    int start = (currentPage - 1) * pageSize + 1;
                    int end = Math.Min(currentPage * pageSize, totalRecords);

                    if (totalRecords == 0)
                    {
                        start = 0;
                        end = 0;
                    }

                    groupBox1.Text = $"Количество записей: {start}-{end} из {totalRecords}";
                    label1.Text = $"Страница {currentPage} из {totalPages}";
                    textBox1.Clear();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных пациентов:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод для маскировки имени и отчества
        private void MaskNameAndPatronymic()
        {
            try
            {
                if (patientTable == null || patientTable.Rows.Count == 0) return;

                foreach (DataRow row in patientTable.Rows)
                {
                    string name = row["Name"]?.ToString();
                    string patronymic = row["Lastname"]?.ToString();
                    string phone = row["Phone_number"]?.ToString();
                    string policy = row["Number_policy"]?.ToString();

                    // Маскируем имя
                    if (!string.IsNullOrEmpty(name) && name.Length > 1)
                    {
                        row["Name"] = $"{name[0]}{new string('*', name.Length - 1)}";
                    }

                    // Маскируем отчество
                    if (!string.IsNullOrEmpty(patronymic) && patronymic.Length > 1)
                    {
                        row["Lastname"] = $"{patronymic[0]}{new string('*', patronymic.Length - 1)}";
                    }

                    // Маскируем номер телефона - показываем только последние 4 цифры
                    if (!string.IsNullOrEmpty(phone) && phone.Length > 5)
                    {
                        row["Phone_number"] = new string('*', phone.Length - 5) + phone.Substring(phone.Length - 5);
                    }

                    // Маскируем номер полиса - показываем только последние 4 цифры
                    if (!string.IsNullOrEmpty(policy) && policy.Length > 4)
                    {
                        row["Number_policy"] = new string('*', policy.Length - 4) + policy.Substring(policy.Length - 4);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при маскировке данных:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Редактирование выбранного пациента
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите запись для редактирования!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int id = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["idPatients"].Value);
                EditPatient ed = new EditPatient(id);
                inactivityTimer.Stop();

                ed.ShowDialog();

                lastActivityTime = DateTime.Now;
                inactivityTimer.Start();

                LoadPatient(); // Обновляем таблицу после редактирования
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при редактировании пациента:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Смена параметра сортировки
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort();
        }

        // Применение фильтра и сортировки к таблице
        private void ApplyFilterAndSort()
        {
            currentPage = 1;
            LoadPatient();
        }

        // Изменение текста поиска
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort();
        }

        // Ввод только чисел для поиска по полису
        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.Numbers(sender, e);
        }

        // Добавление нового пациента
        private void button1_Click(object sender, EventArgs e)
        {
            AddPatient ad = new AddPatient();
            inactivityTimer.Stop();

            ad.ShowDialog();

            lastActivityTime = DateTime.Now;
            inactivityTimer.Start();

            LoadPatient(); // Обновляем таблицу после добавления
        }

        // Выбор пациента для передачи в форму талона
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите пациента!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridViewRow row = dataGridView1.SelectedRows[0];
                int patientId = Convert.ToInt32(row.Cells["idPatients"].Value);

                // Получаем полные данные пациента из базы (без маскировки) и передаём их в талон
                Connect connect = new Connect();
                string connString = connect.ConnectDB();
                string fullName = null;
                using (MySqlConnection con = new MySqlConnection(connString))
                {
                    con.Open();
                    string q = "SELECT Surname, Name, Lastname FROM Patients WHERE idPatients = @id LIMIT 1";
                    using (MySqlCommand cmd = new MySqlCommand(q, con))
                    {
                        cmd.Parameters.AddWithValue("@id", patientId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string surname = reader["Surname"]?.ToString();
                                string name = reader["Name"]?.ToString();
                                string lastname = reader["Lastname"]?.ToString();
                                fullName = $"{surname} {name} {lastname}".Trim();
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(fullName))
                {
                    // Fallback: используем значения из строки таблицы
                    fullName = $"{row.Cells["Surname"].Value} {row.Cells["Name"].Value} {row.Cells["Lastname"].Value}";
                }

                PatientSelected?.Invoke(patientId, fullName);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе пациента:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Сброс фильтров
        private void button5_Click(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = 0;
            textBox5.Text = "";
        }

        // Удаление выбранного пациента
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (selectedId == -1)
                {
                    MessageBox.Show("Выберите запись для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    // Проверка, используется ли пациент в заказах
                    string checkQuery = "SELECT COUNT(*) FROM `Order` WHERE Patients_idPatients = @patientId";
                    using (MySqlCommand checkCmd = new MySqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@patientId", selectedId);
                        int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                        if (count > 0)
                        {
                            MessageBox.Show("Нельзя удалить этого пациента, так как он используется в приемах!",
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    // Подтверждение удаления
                    DialogResult result = MessageBox.Show(
                        "Вы уверены, что хотите удалить запись?",
                        "Подтверждение удаления",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                        return;

                    // Удаление пациента
                    string deleteQuery = "DELETE FROM Patients WHERE idPatients = @id";
                    using (MySqlCommand deleteCmd = new MySqlCommand(deleteQuery, con))
                    {
                        deleteCmd.Parameters.AddWithValue("@id", selectedId);
                        deleteCmd.ExecuteNonQuery();
                        MessageBox.Show("Запись успешно удалена!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                selectedId = -1;
                LoadPatient(); // Обновляем таблицу после удаления
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении пациента:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Получение Id выбранной строки при клике
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                    selectedId = Convert.ToInt32(row.Cells["idPatients"].Value);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе пациента:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void DataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0) return; // Проверяем, что клик не по заголовку

                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                int patientId = Convert.ToInt32(row.Cells["idPatients"].Value);

                // Если уже раскрыт другой пациент — восстанавливаем его в маскированный вид
                if (currentlyUnmaskedPatientId != -1 && currentlyUnmaskedPatientId != patientId)
                {
                    RestoreMaskedRow(currentlyUnmaskedPatientId);
                }

                // Если этот пациент уже раскрыт — ничего не делаем
                if (currentlyUnmaskedPatientId == patientId)
                    return;

                // Загружаем полные данные пациента из БД и запоминаем маскированные значения
                ShowFullPatientData(patientId, row);
                currentlyUnmaskedPatientId = patientId;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при раскрытии данных пациента:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowFullPatientData(int patientId, DataGridViewRow row)
        {
            try
            {
                Connect connect = new Connect();
                string connString = connect.ConnectDB();

                using (MySqlConnection con = new MySqlConnection(connString))
                {
                    con.Open();
                    string query = @"
            SELECT Surname, Name, Lastname, Date_birth, Phone_number, Number_policy 
            FROM Patients 
            WHERE idPatients = @id";

                    using (MySqlCommand cmd = new MySqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@id", patientId);
                        using (MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // Запоминаем маскированные значения (на основе реальных данных)
                                string realName = reader["Name"].ToString();
                                string realLastname = reader["Lastname"].ToString();
                                string realPhone = reader["Phone_number"].ToString();
                                string realPolicy = reader["Number_policy"].ToString();

                                var maskedName = MaskNameValue(realName);
                                var maskedLastname = MaskNameValue(realLastname);
                                var maskedPhone = MaskPhoneValue(realPhone);
                                var maskedPolicy = MaskPolicyValue(realPolicy);

                                maskedBackup[patientId] = (maskedName, maskedLastname, maskedPhone, maskedPolicy);

                                // Обновляем ячейки реальными данными из БД
                                row.Cells["Name"].Value = realName;
                                row.Cells["Lastname"].Value = realLastname;
                                row.Cells["Phone_number"].Value = realPhone;
                                row.Cells["Number_policy"].Value = realPolicy;

                                // Обновляем DataGridView
                                dataGridView1.Refresh();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении данных пациента:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Восстановление маскированной строки по id пациента
        private void RestoreMaskedRow(int patientId)
        {
            try
            {
                // Находим строку с данным id в DataGridView
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    if (r.IsNewRow) continue;
                    if (r.Cells["idPatients"].Value == null) continue;
                    if (Convert.ToInt32(r.Cells["idPatients"].Value) == patientId)
                    {
                        if (maskedBackup.TryGetValue(patientId, out var mask))
                        {
                            r.Cells["Name"].Value = mask.Name;
                            r.Cells["Lastname"].Value = mask.Lastname;
                            r.Cells["Phone_number"].Value = mask.Phone;
                            r.Cells["Number_policy"].Value = mask.Policy;
                        }
                        else
                        {
                            // Если бэкапа нет — применяем простую маску по текущим значениям
                            r.Cells["Name"].Value = MaskNameValue(r.Cells["Name"].Value?.ToString());
                            r.Cells["Lastname"].Value = MaskNameValue(r.Cells["Lastname"].Value?.ToString());
                            r.Cells["Phone_number"].Value = MaskPhoneValue(r.Cells["Phone_number"].Value?.ToString());
                            r.Cells["Number_policy"].Value = MaskPolicyValue(r.Cells["Number_policy"].Value?.ToString());
                        }
                        dataGridView1.Refresh();
                        break;
                    }
                }

                // Сбрасываем отметку
                if (currentlyUnmaskedPatientId == patientId)
                    currentlyUnmaskedPatientId = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при восстановлении маскированной строки:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Утилиты маскировки
        private string MaskNameValue(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;
            return name.Length > 1 ? $"{name[0]}{new string('*', name.Length - 1)}" : name;
        }

        private string MaskPhoneValue(string phone)
        {
            if (string.IsNullOrEmpty(phone)) return phone;
            return phone.Length > 5 ? new string('*', phone.Length - 5) + phone.Substring(phone.Length - 5) : phone;
        }

        private string MaskPolicyValue(string policy)
        {
            if (string.IsNullOrEmpty(policy)) return policy;
            return policy.Length > 4 ? new string('*', policy.Length - 4) + policy.Substring(policy.Length - 4) : policy;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentPage > 1)
                {
                    currentPage--;
                    LoadPatient();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при переходе на предыдущую страницу:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentPage < totalPages)
                {
                    currentPage++;
                    LoadPatient();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при переходе на следующую страницу:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                int page;
                if (int.TryParse(textBox1.Text, out page))
                {
                    if (page >= 1 && page <= totalPages)
                    {
                        currentPage = page;
                        LoadPatient();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при переходе на указанную страницу:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private string BuildFilterSql()
        {
            try
            {
                string where = "WHERE 1=1";

                string search = textBox5.Text.Trim().Replace("'", "''");
                if (!string.IsNullOrEmpty(search))
                {
                    where += $" AND Number_policy LIKE '%{search}%'";
                }

                return where;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при построении фильтра:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "WHERE 1=1"; // Возвращаем базовый фильтр в случае ошибки
            }
        }
        private string GetSortSql()
        {
            try
            {
                if (comboBox2.SelectedIndex == 1)
                    return "ORDER BY Surname ASC";
                else if (comboBox2.SelectedIndex == 2)
                    return "ORDER BY Surname DESC";

                return "ORDER BY idPatients DESC";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при построении сортировки:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "ORDER BY idPatients DESC"; // Возвращаем базовую сортировку в случае ошибки
            }
        }
        private int GetTotalCount(string filterSql)
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string query = $"SELECT COUNT(*) FROM Patients {filterSql}";
                    MySqlCommand cmd = new MySqlCommand(query, con);
                    return Convert.ToInt32(cmd.ExecuteScalar());
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Ошибка при подсчете общего количества записей:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0; // Возвращаем 0 в случае ошибки
            }
        }
        private void RegisterActivityHandlers(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                ctrl.MouseMove += ResetActivity;
                ctrl.MouseClick += ResetActivity;
                ctrl.KeyDown += ResetActivity;

                // Рекурсивно для вложенных контролов
                if (ctrl.HasChildren)
                    RegisterActivityHandlers(ctrl);
            }
        }
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            inactivityTimer.Stop();
            inactivityTimer.Tick -= InactivityTimer_Tick;
            inactivityTimer.Dispose();

            OnSessionExpired = null;

            base.OnFormClosed(e);
        }
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);

            if (this.Visible)
            {
                lastActivityTime = DateTime.Now;
                inactivityTimer.Start();
            }
            else
            {
                inactivityTimer.Stop();
            }
        }
    }
}
