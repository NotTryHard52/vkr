using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class UchetTalona : Form
    {
        // Строка подключения к базе данных
        string connectionString;

        // Таблица для хранения данных о талонах
        DataTable orderTable;
        int currentPage = 1;
        int pageSize = 10;
        int totalRecords = 0;
        int totalPages = 1;
        bool isGlav = false;
        private Timer inactivityTimer;
        private DateTime lastActivityTime;
        private const int timeoutSeconds = 120;
        public event Action OnSessionExpired;

        public UchetTalona(bool isGlav = false)
        {
            try
            {
                InitializeComponent();
                this.isGlav = isGlav;
                dataGridView1.SizeChanged += (s, e) => ReloadOrderTable();
                if (isGlav)
                {
                    button6.Visible = true;
                }

                inactivityTimer = new Timer();
                inactivityTimer.Interval = 1000; // проверка каждую секунду
                inactivityTimer.Tick += InactivityTimer_Tick;
                inactivityTimer.Start();

                lastActivityTime = DateTime.Now;

                // отслеживание активности
                RegisterActivityHandlers(this);
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Ошибка при инициализации формы:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResetActivity(object sender, EventArgs e)
        {
            lastActivityTime = DateTime.Now;
        }

        private void InactivityTimer_Tick(object sender, EventArgs e)
        {
            if ((DateTime.Now - lastActivityTime).TotalSeconds >= timeoutSeconds)
            {
                inactivityTimer.Stop();
                OnSessionExpired?.Invoke();
            }
        }

        private void RegisterActivityHandlers(Control parent)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при регистрации обработчиков активности:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при расчёте размера страницы:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 10; // возвращаем значение по умолчанию
            }
        }

        private int GetTotalCount(string filterSql)
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    string query = "SELECT COUNT(*) FROM `Order` o " +
                                   "JOIN StatusesPriem st ON o.Status = st.idStatusesPriem " +
                                   filterSql;

                    MySqlCommand cmd = new MySqlCommand(query, con);
                    return Convert.ToInt32(cmd.ExecuteScalar());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении общего количества записей:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0; // возвращаем 0 в случае ошибки
            }
        }

        private string BuildFilterSql()
        {
            try
            {
                string where = "WHERE 1=1";

                string status = comboBox1.SelectedItem?.ToString();
                if (!string.IsNullOrEmpty(status) && status != "Все статусы")
                {
                    where += $" AND st.name = '{status.Replace("'", "''")}'";
                }

                string search = textBox5.Text.Trim();
                if (!string.IsNullOrEmpty(search))
                {
                    where += $" AND o.idOrder LIKE '%{search}%'";
                }

                return where;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при построении SQL фильтра:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "WHERE 1=1"; // возвращаем базовый фильтр в случае ошибки
            }
        }
        private string GetSortSql()
        {
            try
            {
                if (comboBox2.SelectedIndex == 1)
                    return "ORDER BY sc.date ASC";
                else if (comboBox2.SelectedIndex == 2)
                    return "ORDER BY sc.date DESC";

                return " ";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при построении SQL сортировки:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return " "; // возвращаем пустую строку в случае ошибки
            }
        }

        // Событие загрузки формы
        private void UchetTalona_Load(object sender, EventArgs e)
        {
            try
            {
                FillStatus();          // Заполнение comboBox статусами
                comboBox1.SelectedIndex = 0; // По умолчанию "Все"
                comboBox2.SelectedIndex = 0; // По умолчанию без сортировки
                this.Shown += (s, ev) => ReloadOrderTable();
                if (!isGlav)
                {
                    label2.Visible = false;
                    button6.Visible = false;
                }
                dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
                dataGridView1.RowTemplate.Height = 40;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке формы:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Подсчёт общей выручки (не учитываются отменённые и созданные талоны)
        private void UpdateRevenueSum()
        {
            try
            {
                Connect connect = new Connect();
                connectionString = connect.ConnectDB();

                string filterSql = BuildFilterSql();

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    string query = @"
                        SELECT IFNULL(SUM(o.TotalSum), 0)
                        FROM `Order` o
                        JOIN StatusesPriem st ON o.Status = st.idStatusesPriem
                        WHERE st.name NOT IN ('Отменён', 'Создан');
                        ";

                    MySqlCommand cmd = new MySqlCommand(query, con);

                    object result = cmd.ExecuteScalar();

                    decimal total = 0;

                    if (result != null && result != DBNull.Value)
                        total = Convert.ToDecimal(result);

                    label2.Text = $"Общая выручка: {total:N2} руб.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при подсчёте общей выручки:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Заполнение comboBox1 статусами из базы
        private void FillStatus()
        {
            try
            {
                Connect connect = new Connect();
                connectionString = connect.ConnectDB();
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT DISTINCT name FROM StatusesPriem;";
                    MySqlCommand cmd = new MySqlCommand(query, con);
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        comboBox1.Items.Clear();
                        comboBox1.Items.Add("Все статусы"); // Добавляем опцию "Все статусы"
                        while (reader.Read())
                        {
                            string status = reader["name"].ToString();
                            comboBox1.Items.Add(status); // Добавляем каждый статус
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении статусов:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Событие изменения выбора в comboBox1
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort(); // Применяем фильтр и сортировку
        }

        // Применение фильтров и сортировки к DataGridView
        private void ApplyFilterAndSort()
        {
            // Переключаем на серверную фильтрацию/сортировку и сбрасываем на первую страницу
            currentPage = 1;
            ReloadOrderTable();
        }

        // Событие изменения выбора сортировки
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort();
        }

        // Событие изменения текста поиска
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort();
        }

        // Кнопка "Открыть талон"
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    int orderId = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["Номер"].Value);
                    if (!isGlav)
                    {
                        ViewPriem v = new ViewPriem(orderId, false);
                        inactivityTimer.Stop();

                        var result = v.ShowDialog();

                        lastActivityTime = DateTime.Now;
                        inactivityTimer.Start(); // открываем форму просмотра талона
                    }
                    else
                    {
                        ViewPriem v = new ViewPriem(orderId, true);
                        inactivityTimer.Stop();

                        var result = v.ShowDialog();

                        lastActivityTime = DateTime.Now;
                        inactivityTimer.Start(); // открываем форму просмотра талона
                    }
                    ReloadOrderTable(); // обновляем таблицу после просмотра
                }
                else
                {
                    MessageBox.Show("Пожалуйста, выберите талон.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии талона:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Загрузка данных о талонах из базы
        private void ReloadOrderTable()
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

                    // защита от выхода за границы
                    if (currentPage > totalPages)
                        currentPage = totalPages == 0 ? 1 : totalPages;

                    int offset = (currentPage - 1) * pageSize;

                    string query = $@"
                    SELECT 
                        o.idOrder AS 'Номер',
                        CONCAT(d.surname, ' ', d.name, ' ', d.lastname) AS 'Врач',
                        DATE_FORMAT(sc.date, '%d.%m.%Y') AS 'Дата',
                        DATE_FORMAT(sc.time, '%H:%i') AS 'Время',
                        CONCAT(r.surname, ' ', r.name, ' ', r.lastname) AS 'Регистратор',
                        CONCAT(p.surname, ' ', p.name, ' ', p.lastname) AS 'Пациент',
                        st.name AS 'Статус',
                        o.sum AS 'Сумма',
                        o.Discount AS 'Скидка',
                        o.TotalSum AS 'Итого'
                    FROM `Order` o
                    JOIN Schedule sc ON o.Schedule = sc.idSchedule
                    JOIN Doctors d ON sc.idDoctor = d.idDoctors
                    JOIN `Users` r ON o.User = r.idUsers
                    JOIN Patients p ON o.Patients_idPatients = p.idPatients
                    JOIN StatusesPriem st ON o.Status = st.idStatusesPriem
                    {filterSql}
                    {sortSql}
                    LIMIT @offset, @pageSize;
                ";

                    MySqlCommand cmd = new MySqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@offset", offset);
                    cmd.Parameters.AddWithValue("@pageSize", pageSize);

                    DataTable table = new DataTable();
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    da.Fill(table);

                    orderTable = table;

                    dataGridView1.DataSource = table;
                    dataGridView1.Columns["Сумма"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns["Итого"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns["Скидка"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns["Время"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView1.Columns["Статус"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView1.Columns["Врач"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    dataGridView1.Columns["Пациент"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    dataGridView1.Columns["Регистратор"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    dataGridView1.Columns["Номер"].Width = 75;
                    dataGridView1.Columns["Сумма"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dataGridView1.Columns["Скидка"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dataGridView1.Columns["Итого"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dataGridView1.Columns["Дата"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dataGridView1.Columns["Время"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dataGridView1.Columns["Статус"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                    int start = (currentPage - 1) * pageSize + 1;
                    int end = Math.Min(currentPage * pageSize, totalRecords);

                    if (totalRecords == 0)
                    {
                        start = 0;
                        end = 0;
                    }

                    groupBox2.Text = $"Количество записей: {start}-{end} из {totalRecords}";
                    label1.Text = $"Страница {currentPage} из {totalPages}";
                    textBox1.Clear();
                    UpdateRevenueSum();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных о талонах:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Кнопка "Сброс фильтров"
        private void button5_Click(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = 0;
            comboBox1.SelectedIndex = 0;
            textBox5.Text = "";
        }

        // Изменение цвета строк в зависимости от статуса
        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            try
            {
                if (dataGridView1.Rows[e.RowIndex].Cells["Статус"].Value == null)
                    return;

                string status = dataGridView1.Rows[e.RowIndex].Cells["Статус"].Value.ToString().ToLower();

                if (status.Contains("заверш"))
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightGreen; // завершён
                }
                else if (status.Contains("отмен"))
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightCoral; // отменён
                }
                else if (status.Contains("создан"))
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightYellow; // создан
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White; // прочие
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Ошибка при раскраске строк:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.Numbers(sender, e);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentPage < totalPages)
                {
                    currentPage++;
                    ReloadOrderTable();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при переходе на следующую страницу:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentPage > 1)
                {
                    currentPage--;
                    ReloadOrderTable();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при переходе на предыдущую страницу:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                int page;
                if (int.TryParse(textBox1.Text, out page))
                {
                    if (page >= 1 && page <= totalPages)
                    {
                        currentPage = page;
                        ReloadOrderTable();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при переходе на указанную страницу:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.Numbers(sender, e);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OtchetExcel v = new OtchetExcel();

            inactivityTimer.Stop();

            v.ShowDialog();

            lastActivityTime = DateTime.Now;
            inactivityTimer.Start();
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
