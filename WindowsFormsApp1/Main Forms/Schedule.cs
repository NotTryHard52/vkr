using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Schedule : Form
    {
        // Строка подключения к базе данных
        string connectionString;
        // Таблица с данными расписания
        DataTable scheduleTable;
        // Таблица с данными врачей для поиска
        DataTable doctorsTable;
        // Флаг, чтобы предотвратить повторный вход в обработчик фильтрации
        bool isFilteringCombo = false;
        // ID выбранной записи расписания
        int selectedScheduleId = -1;
        // Событие, передающее выбранную запись (для формы талона)
        public event Action<int, string, string, string> ScheduleSelected;
        // Статус выбранного приёма ("Свободно", "Занято")
        string status;
        // Флаг: открыто из формы талона
        private bool openedFromTalon = false;
        // Текущая системная дата для генерации кнопок дат (2 недели)
        private DateTime selectedDate = DateTime.Today;
        int currentPage = 1;
        int pageSize = 10;
        int totalRecords = 0;
        int totalPages = 1;
        private Timer resizeTimer = new Timer();
        private bool loadingSchedule = false;

        // Конструктор формы
        public Schedule(bool fromTalon = false)
        {
            InitializeComponent();
            openedFromTalon = fromTalon;
            // Кнопка выбора записи видна только если форма открыта из талона
            if(openedFromTalon)
            {
                button1.Visible = false; // Скрываем кнопку "Добавить"
                button2.Visible = false;
                button3.Visible = false;
            }
            this.Resize += Schedule_Resize;
            resizeTimer.Interval = 250;
            resizeTimer.Tick += (s, e) =>
            {
                resizeTimer.Stop();
                ReloadOrderTable(); 
            };

            this.Resize += (s, e) =>
            {
                UpdateScheduleCardLayout();
                resizeTimer.Stop();
                resizeTimer.Start();
            };
        }

        private void FillDoctors()
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    string query = @"
            SELECT 
                idDoctors, 
                CONCAT(Surname, ' ', Name, ' ', Lastname) AS FullName
            FROM Doctors
            ORDER BY Surname";

                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, con);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    // Сохраняем оригинальную таблицу для фильтрации
                    doctorsTable = dt;

                    // Разрешаем ввод в combobox и настраиваем поиск
                    comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
                    comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;

                    var ac = new AutoCompleteStringCollection();
                    foreach (DataRow r in dt.Rows)
                        ac.Add(r["FullName"].ToString());
                    comboBox1.AutoCompleteCustomSource = ac;

                    // Привязываем полную таблицу
                    comboBox1.DataSource = doctorsTable.Copy();
                    comboBox1.DisplayMember = "FullName";   // что показывать
                    comboBox1.ValueMember = "idDoctors";    // значение (ID)
                    comboBox1.SelectedIndex = -1;           // ничего не выбрано по умолчанию

                    // Подписываемся на изменение текста для фильтрации
                    comboBox1.TextChanged -= ComboBox1_TextChanged;
                    comboBox1.TextChanged += ComboBox1_TextChanged;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке врачей:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (isFilteringCombo) return;
                if (doctorsTable == null) return;

                string txt = comboBox1.Text.Trim();

                int selStart = comboBox1.SelectionStart;
                string current = comboBox1.Text;

                try
                {
                    isFilteringCombo = true;
                    comboBox1.BeginUpdate();

                    if (string.IsNullOrEmpty(txt))
                    {
                        comboBox1.DataSource = doctorsTable.Copy();
                    }
                    else
                    {
                        DataView dv = new DataView(doctorsTable);
                        string filter = txt.Replace("'", "''");
                        dv.RowFilter = string.Format("FullName LIKE '%{0}%'", filter);
                        DataTable filtered = dv.ToTable();
                        comboBox1.DataSource = filtered;
                    }

                    // Восстанавливаем текст и позицию курсора
                    comboBox1.Text = current;
                    try { comboBox1.SelectionStart = Math.Min(selStart, comboBox1.Text.Length); } catch { }

                    // Показываем список подсказок только если есть результаты
                    try
                    {
                        if (comboBox1.Focused && comboBox1.Items.Count > 1)
                        {
                            comboBox1.DroppedDown = true;
                        }
                    }
                    catch { }
                }
                finally
                {
                    try { comboBox1.EndUpdate(); } catch { }
                    isFilteringCombo = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка фильтрации врачей:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Schedule_Resize(object sender, EventArgs e)
        {
            UpdateScheduleCardLayout();
        }

        private void GenerateDates()
        {
            try
            {
                flowLayoutPanel1.Controls.Clear();

                DateTime start = DateTime.Today;

                for (int i = 0; i < 14; i++)
                {
                    DateTime date = start.AddDays(i);

                    Button btn = new Button();
                    btn.Width = 100;
                    btn.Height = 60;
                    btn.Tag = date;

                    btn.Text = $"{date:dd.MM}\n{GetDayName(date)}";

                    btn.Click += DateButton_Click;

                    StyleDateButton(btn, date == selectedDate);

                    flowLayoutPanel1.Controls.Add(btn);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка генерации кнопок дат:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetDayName(DateTime date)
        {
            return date.ToString("ddd"); // пн, вт, ср
        }

        private void DateButton_Click(object sender, EventArgs e)
        {
            try
            {
                Button clicked = sender as Button;
                selectedDate = (DateTime)clicked.Tag;

                // Перерисовать кнопки (подсветка)
                foreach (Button btn in flowLayoutPanel1.Controls)
                {
                    StyleDateButton(btn, (DateTime)btn.Tag == selectedDate);
                }

                LoadScheduleByDate(selectedDate);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе даты:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StyleDateButton(Button btn, bool isActive)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;

            if (isActive)
            {
                btn.BackColor = Color.FromArgb(91, 122, 196);
                btn.ForeColor = Color.White;
            }
            else
            {
                btn.BackColor = Color.White;
                btn.ForeColor = Color.Black;
            }
        }

        // Загрузка формы
        private void Schedule_Load(object sender, EventArgs e)
        {
            Connect connect = new Connect();
            connectionString = connect.ConnectDB(); 
            
            FillStatuses();
            FillDoctors();
            comboBox2.SelectedIndex = 0;
            GenerateDates();
            LoadScheduleByDate(DateTime.Today);
            InputLimit.DateOrder(dateTimePicker1);
        }

        private void LoadScheduleByDate(DateTime date)
        {
            try
            {
                pageSize = CalculatePageSize();

                string selectedStatus = comboBox2.SelectedItem?.ToString();

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    string countQuery = @"
                        SELECT COUNT(*)
                        FROM Schedule s
                        JOIN Statuses st ON s.Status = st.idStatuses
                        WHERE s.Date = @date
                        ";

                    if (!string.IsNullOrEmpty(selectedStatus) && selectedStatus != "Все")
                    {
                        countQuery += " AND st.statusname = @status ";
                    }

                    MySqlCommand countCmd = new MySqlCommand(countQuery, con);
                    countCmd.Parameters.AddWithValue("@date", date.ToString("yyyy-MM-dd"));

                    if (!string.IsNullOrEmpty(selectedStatus) && selectedStatus != "Все")
                    {
                        countCmd.Parameters.AddWithValue("@status", selectedStatus);
                    }

                    totalRecords = Convert.ToInt32(countCmd.ExecuteScalar());
                    totalPages = Math.Max(1, (int)Math.Ceiling(totalRecords / (double)pageSize));

                    if (currentPage > totalPages)
                        currentPage = totalPages;

                    if (currentPage < 1)
                        currentPage = 1;

                    int offset = (currentPage - 1) * pageSize;

                    string query = @"
                        SELECT
                            s.idSchedule, 
                            CONCAT(d.Surname, ' ', d.Name, ' ', d.Lastname) AS Doctor,
                            TIME_FORMAT(s.Time, '%H:%i') as Time,
                            st.statusname as Status
                        FROM Schedule s
                        JOIN Doctors d ON s.idDoctor = d.idDoctors
                        JOIN Statuses st ON s.Status = st.idStatuses
                        WHERE s.Date = @date
                        ";

                    if (!string.IsNullOrEmpty(selectedStatus) && selectedStatus != "Все")
                    {
                        query += " AND st.statusname = @status ";
                    }

                    query += " ORDER BY s.Time LIMIT @limit OFFSET @offset";

                    MySqlCommand cmd = new MySqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@date", date.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@limit", pageSize);
                    cmd.Parameters.AddWithValue("@offset", offset);

                    if (!string.IsNullOrEmpty(selectedStatus) && selectedStatus != "Все")
                    {
                        cmd.Parameters.AddWithValue("@status", selectedStatus);
                    }

                    List<ScheduleItem> list = new List<ScheduleItem>();

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(new ScheduleItem
                            {
                                Id = reader.GetInt32("idSchedule"),
                                DoctorName = reader.GetString("Doctor"),
                                Time = reader.GetString("Time"),
                                Status = reader.GetString("Status")
                            });
                        }
                    }

                    int start = (currentPage - 1) * pageSize + 1;
                    int end = Math.Min(currentPage * pageSize, totalRecords);

                    if (totalRecords == 0)
                    {
                        start = 0;
                        end = 0;
                    }

                    groupBox2.Text = $"Количество записей: {start}-{end} из {totalRecords}";

                    ShowSchedule(list);
                    UpdatePaginationUI();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке расписания:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdatePaginationUI()
        {
            label2.Text = $"Страница {currentPage} из {totalPages}";
        }

        private void ReloadOrderTable()
        {
            LoadScheduleByDate(selectedDate);
        }

        private Control CreateScheduleCard(ScheduleItem item)
        {
            try
            {
                var card = new ScheduleCard(item);

                card.CardClicked += Card_Click;

                card.Width = 600;
                card.Height = 80;
                card.Margin = new Padding(8);

                return card;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании карточки расписания:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new Label { Text = "Ошибка создания карточки" };
            }
        }

        private void ShowSchedule(List<ScheduleItem> data)
        {
            try
            {
                flowLayoutPanel2.Controls.Clear();

                foreach (var item in data)
                {
                    var card = CreateScheduleCard(item);
                    flowLayoutPanel2.Controls.Add(card);
                }
                UpdateScheduleCardLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отображении расписания:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public class ScheduleItem
        {
            public int Id { get; set; }
            public string DoctorName { get; set; }
            public string Time { get; set; }
            public string Status { get; set; }
        }

        // Метод для заполнения ComboBox2 статусами приёмов
        private void FillStatuses()
        {
            try
            {
                Connect connect = new Connect();
                connectionString = connect.ConnectDB();
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT DISTINCT statusname FROM Statuses;";
                    MySqlCommand cmd = new MySqlCommand(query, con);
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        comboBox2.Items.Clear();
                        comboBox2.Items.Add("Все"); // Для отображения всех записей
                        while (reader.Read())
                        {
                            string statuses = reader["statusname"].ToString();
                            comboBox2.Items.Add(statuses);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке статусов:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Обработчик изменения статуса в ComboBox2
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            currentPage = 1;
            LoadScheduleByDate(selectedDate);
        }

        // Добавление новой записи расписания
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.SelectedValue == null)
                {
                    MessageBox.Show("Выберите врача!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int doctorId = Convert.ToInt32(comboBox1.SelectedValue);
                DateTime date = dateTimePicker1.Value.Date;
                TimeSpan time = dateTimePicker2.Value.TimeOfDay;

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    // Проверка на существующую запись у этого врача в это время
                    string checkQuery = @"
                    SELECT COUNT(*) FROM Schedule
                    WHERE idDoctor = @idDoctor AND date = @date AND time = @time";
                    using (MySqlCommand checkCmd = new MySqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@idDoctor", doctorId);
                        checkCmd.Parameters.AddWithValue("@date", date);
                        checkCmd.Parameters.AddWithValue("@time", time);

                        int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                        if (count > 0)
                        {
                            MessageBox.Show("У этого врача уже есть запись на указанное время!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    // Получаем ID статуса "Свободно"
                    int statusId = 0;
                    string getStatusQuery = "SELECT idStatuses FROM Statuses WHERE statusname = 'Свободно' LIMIT 1;";
                    using (MySqlCommand statusCmd = new MySqlCommand(getStatusQuery, con))
                    {
                        object result = statusCmd.ExecuteScalar();
                        if (result == null)
                        {
                            MessageBox.Show("Не найден статус 'Свободно' в таблице Statuses!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        statusId = Convert.ToInt32(result);
                    }

                    // Вставка новой записи
                    string insertQuery = @"
                    INSERT INTO Schedule (idDoctor, date, time, Status)
                    VALUES (@idDoctor, @date, @time, @status)";
                    using (MySqlCommand insertCmd = new MySqlCommand(insertQuery, con))
                    {
                        insertCmd.Parameters.AddWithValue("@idDoctor", doctorId);
                        insertCmd.Parameters.AddWithValue("@date", date);
                        insertCmd.Parameters.AddWithValue("@time", time);
                        insertCmd.Parameters.AddWithValue("@status", statusId);
                        insertCmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Запись успешно добавлена!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    comboBox1.SelectedIndex = -1;
                }

                LoadScheduleByDate(selectedDate);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении записи:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // Обработчик изменения времени в dateTimePicker2
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                // Если выбранная дата сегодня, проверяем, чтобы время не было в прошлом
                if (loadingSchedule)
                    return;

                if (dateTimePicker1.Value.Date == DateTime.Now.Date)
                {
                    DateTime selectedDateTime =
                        dateTimePicker1.Value.Date + dateTimePicker2.Value.TimeOfDay;

                    if (selectedDateTime < DateTime.Now)
                    {
                        dateTimePicker2.Value = DateTime.Now;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при изменении времени:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetDoctorByName(string fullName)
        {
            try
            {
                comboBox1.DataSource = doctorsTable.Copy();
                comboBox1.DisplayMember = "FullName";
                comboBox1.ValueMember = "idDoctors";

                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    DataRowView row = (DataRowView)comboBox1.Items[i];

                    if (row["FullName"].ToString() == fullName)
                    {
                        comboBox1.SelectedIndex = i;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при установке врача по имени:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Card_Click(ScheduleItem item)
        {
            try
            {
                loadingSchedule = true;
                // сохраняем выбранную запись
                selectedScheduleId = item.Id;
                status = item.Status;

                // заполняем дату
                dateTimePicker1.Value = selectedDate;

                // заполняем время
                dateTimePicker2.Value = DateTime.Today.Add(TimeSpan.Parse(item.Time));
                loadingSchedule = false;

                // устанавливаем врача
                SetDoctorByName(item.DoctorName);


                // если открыто из талона
                if (openedFromTalon)
                {
                    if (item.Status != "Свободно")
                    {
                        MessageBox.Show("Этот слот уже занят!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    DateTime dateTime = selectedDate.Date + TimeSpan.Parse(item.Time);

                    if (dateTime < DateTime.Now)
                    {
                        MessageBox.Show("Нельзя выбрать прошлое время!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    ScheduleSelected?.Invoke(
                        item.Id,
                        item.DoctorName,
                        selectedDate.ToString("yyyy-MM-dd"),
                        item.Time
                    );

                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе записи:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Обработчик кнопки "Сброс фильтра статуса"
        private void button5_Click(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = 0; // Выбираем "Все"
        }

        // Редактирование выбранной записи расписания
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что запись свободна
                if (status != "Свободно")
                {
                    MessageBox.Show("Редактирование невозможно — расписание уже занято!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    selectedScheduleId = -1;
                    return;
                }

                if (selectedScheduleId == -1)
                {
                    MessageBox.Show("Выберите запись для редактирования!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (comboBox1.SelectedValue == null)
                {
                    MessageBox.Show("Выберите врача!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int doctorId = Convert.ToInt32(comboBox1.SelectedValue);
                DateTime date = dateTimePicker1.Value.Date;
                TimeSpan time = dateTimePicker2.Value.TimeOfDay;

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    // Проверка, что у врача нет другой записи в это время
                    string checkQuery = @"
                    SELECT COUNT(*) FROM Schedule
                    WHERE idDoctor = @idDoctor AND date = @date AND time = @time 
                          AND idSchedule <> @idSchedule";
                    using (MySqlCommand checkCmd = new MySqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@idDoctor", doctorId);
                        checkCmd.Parameters.AddWithValue("@date", date);
                        checkCmd.Parameters.AddWithValue("@time", time);
                        checkCmd.Parameters.AddWithValue("@idSchedule", selectedScheduleId);

                        int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                        if (count > 0)
                        {
                            MessageBox.Show("У этого врача уже есть запись на указанное время!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    // Обновление записи в базе
                    string updateQuery = @"
                    UPDATE Schedule
                    SET idDoctor = @doctorId, date = @date, time = @time
                    WHERE idSchedule = @idSchedule";
                    using (MySqlCommand updateCmd = new MySqlCommand(updateQuery, con))
                    {
                        updateCmd.Parameters.AddWithValue("@doctorId", doctorId);
                        updateCmd.Parameters.AddWithValue("@date", date);
                        updateCmd.Parameters.AddWithValue("@time", time);
                        updateCmd.Parameters.AddWithValue("@idSchedule", selectedScheduleId);
                        updateCmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Запись успешно обновлена!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                selectedScheduleId = -1;
                LoadScheduleByDate(selectedDate);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при редактировании записи:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Удаление выбранной записи расписания
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (selectedScheduleId == -1)
                {
                    MessageBox.Show("Выберите запись для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (status != "Свободно")
                {
                    MessageBox.Show("Нельзя удалить занятую запись расписания!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    "Вы уверены, что хотите удалить выбранную запись расписания?",
                    "Подтверждение удаления",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.No) return;

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string deleteQuery = "DELETE FROM Schedule WHERE idSchedule = @id";
                    using (MySqlCommand deleteCmd = new MySqlCommand(deleteQuery, con))
                    {
                        deleteCmd.Parameters.AddWithValue("@id", selectedScheduleId);
                        deleteCmd.ExecuteNonQuery();
                        MessageBox.Show("Запись успешно удалена!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                selectedScheduleId = -1;
                LoadScheduleByDate(selectedDate); // Обновляем расписание
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении записи:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void UpdateScheduleCardLayout()
        {
            try
            {
                int panelWidth = flowLayoutPanel2.ClientSize.Width;

                int cardWidth;

                if (panelWidth < 700)
                    cardWidth = panelWidth - 25;   // 1 колонка
                else if (panelWidth < 1200)
                    cardWidth = (panelWidth / 2) - 25; // 2 колонки
                else
                    cardWidth = (panelWidth / 3) - 25; // 3 колонки

                foreach (Control ctrl in flowLayoutPanel2.Controls)
                {
                    ctrl.Width = cardWidth;
                    ctrl.Height = 90; // можно под себя
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении макета карточек:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button7_Click(object sender, EventArgs e)
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

        private void button8_Click(object sender, EventArgs e)
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
                        ReloadOrderTable();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при переходе на указанную страницу:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private int CalculatePageSize()
        {
            try
            {
                int cardHeight = 98; // 90 + margin
                int availableHeight = flowLayoutPanel2.ClientSize.Height;

                int rows = availableHeight / cardHeight;

                int columns = 1;

                int panelWidth = flowLayoutPanel2.ClientSize.Width;

                if (panelWidth >= 1200)
                    columns = 3;
                else if (panelWidth >= 700)
                    columns = 2;

                return Math.Max(1, rows * columns);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при расчете размера страницы:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1; // Возвращаем минимальный размер страницы в случае ошибки
            }
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.RussianComboBox(sender, e);
        }
    }
}