using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    public partial class Talon : Form
    {
        // Строка подключения к базе данных
        private string connectionString;

        // Информация о текущем пользователе
        private string currentUserFullName;
        private int currentUserId;         // id текущего пользователя

        // Выбранные элементы формы
        private int selectedScheduleId;    // id выбранного расписания
        private int selectedPatientId;     // id выбранного пациента

        // Конструктор формы
        public Talon(string userFullName, int userId)
        {
            InitializeComponent();
            currentUserFullName = userFullName;  // Сохраняем имя текущего пользователя
            currentUserId = userId;              // Сохраняем id текущего пользователя
        }

        // Событие загрузки формы
        private void Talon_Load(object sender, EventArgs e)
        {
            try
            {
                // Отображение текущего пользователя
                label3.Text = "Пользователь: " + currentUserFullName;

                // Подключение к базе данных
                Connect connect = new Connect();
                connectionString = connect.ConnectDB();

                // Загрузка услуг из базы данных
                LoadServices();

                // Настройка dataGridView2 для выбранных услуг
                dataGridView2.Columns.Clear();
                dataGridView2.Columns.Add("ServiceName", "Наименование");
                dataGridView2.Columns.Add("Price", "Цена");
                dataGridView2.Columns.Add("ServiceId", "ServiceId"); // скрытый столбец для id услуги
                dataGridView2.Columns["ServiceId"].Visible = false;
                dataGridView2.Columns["ServiceName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns["Price"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridView2.AllowUserToAddRows = false;

                // Добавляем эффект подсветки при наведении на строки
                var hoverEffect = new HoverDataGridView(dataGridView1);
                var hoverEffect2 = new HoverDataGridView(dataGridView2);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке формы: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Загрузка услуг из базы данных в dataGridView1
        private void LoadServices()
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    DataTable t = new DataTable();
                    MySqlCommand cmd = new MySqlCommand(@"
                    SELECT s.idServices AS ServiceId,
                           s.Name AS ServiceName,
                           s.Price
                    FROM Services s;", con);
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    da.Fill(t);

                    // Настройка DataGridView
                    dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect; // выделение всей строки
                    dataGridView1.DataSource = t; // привязка данных
                    dataGridView1.Columns["ServiceId"].Visible = false; // скрываем id
                    dataGridView1.Columns["ServiceName"].HeaderText = "Наименование";
                    dataGridView1.Columns["Price"].HeaderText = "Цена";
                    dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке услуг: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Выбор расписания приёмов
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                Schedule scheduleForm = new Schedule(true); // передаём true для выбора

                scheduleForm.ScheduleSelected += (scheduleId, doctorName, date, time) =>
                {
                    try
                    {
                        selectedScheduleId = scheduleId;

                        label1.Text = "Врач: " + doctorName;
                        label7.Text = "Дата приема: " + date;
                        label8.Text = "Время приема: " + time;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                };
                scheduleForm.ShowDialog(); // показываем форму выбора
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при выборе расписания: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Выбор пациента
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                Patient patientForm = new Patient(true); // передаём true для выбора
                patientForm.PatientSelected += (patientId, fullName) =>
                {
                    // Получаем данные выбранного пациента
                    selectedPatientId = patientId;
                    label9.Text = "Пациент: " + fullName;
                };
                patientForm.ShowDialog(); // показываем форму выбора
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при выборе пациента: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Добавление услуги в талон
        private void button1_Click(object sender, EventArgs e)
        {
            AddSelectedService();
        }

        // Подсчёт суммы и скидки
        private void UpdateTotal()
        {
            try
            {
                decimal total = 0;

                // Суммируем цену всех выбранных услуг
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (decimal.TryParse(row.Cells["Price"].Value.ToString(), out decimal price))
                        total += price;
                }

                decimal discount = 0;

                // Применяем скидку 5% если сумма больше 5000
                if (total >= 5000)
                {
                    discount = total * 0.05m;
                }

                decimal finalTotal = total - discount;

                // Отображение итогов на форме
                label6.Text = $"Итого: {total:N2} руб.";
                label5.Text = $"Скидка: {discount:N2} руб.";
                label10.Text = $"К оплате: {finalTotal:N2} руб.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при подсчёте суммы: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Сохранение талона в базу и возможность печати
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверка выбора пациента и расписания
                if (selectedPatientId == 0 || selectedScheduleId == 0)
                {
                    MessageBox.Show("Выберите пациента и расписание!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Проверяем, нет ли у пациента записи на этот же день
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    string checkSql = @"
        SELECT COUNT(*) 
        FROM `Order` o
        JOIN Schedule s ON o.schedule = s.idSchedule
        WHERE o.Patients_idPatients = @patientId
          AND DATE(s.Date) = (
                SELECT DATE(Date) 
                FROM Schedule 
                WHERE idSchedule = @currentSchedule
          )
          AND o.Status <> 0;";

                    using (MySqlCommand cmd = new MySqlCommand(checkSql, con))
                    {
                        cmd.Parameters.AddWithValue("@patientId", selectedPatientId);
                        cmd.Parameters.AddWithValue("@currentSchedule", selectedScheduleId);

                        int count = Convert.ToInt32(cmd.ExecuteScalar());

                        if (count > 0)
                        {
                            MessageBox.Show(
                                "У этого пациента уже есть запись на выбранную дату!",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);

                            return; // стоп — талон не создаём
                        }
                    }
                }

                // Расчёт общей суммы и скидки
                decimal total = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    total += Convert.ToDecimal(row.Cells["Price"].Value);
                }

                decimal discount = 0;
                if (total > 5000)
                {
                    discount = total * 0.05m; // 5% скидка
                }

                decimal finalTotal = total - discount;

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    using (MySqlTransaction transaction = con.BeginTransaction()) // Используем транзакцию для целостности данных
                    {
                        try
                        {
                            int orderId;

                            // Вставляем запись в Order
                            string insertOrder = @"
                    INSERT INTO `Order` (sum, Discount, TotalSum, schedule, Patients_idPatients, Status, User)
                    VALUES (@sum, @discount, @totalSum, @schedule, @patientId, @status, @userId);
                    SELECT LAST_INSERT_ID();";

                            using (MySqlCommand cmd = new MySqlCommand(insertOrder, con, transaction))
                            {
                                cmd.Parameters.AddWithValue("@sum", total);
                                cmd.Parameters.AddWithValue("@discount", discount);
                                cmd.Parameters.AddWithValue("@totalSum", finalTotal);
                                cmd.Parameters.AddWithValue("@schedule", selectedScheduleId);
                                cmd.Parameters.AddWithValue("@patientId", selectedPatientId);
                                cmd.Parameters.AddWithValue("@status", 3); // 3 = Создан
                                cmd.Parameters.AddWithValue("@userId", currentUserId);

                                orderId = Convert.ToInt32(cmd.ExecuteScalar());
                            }

                            // Вставляем выбранные услуги в OrderServices
                            string insertService = @"
                    INSERT INTO OrderServices (OrderId, ServicesId)
                    VALUES (@orderId, @serviceId)";
                            foreach (DataGridViewRow row in dataGridView2.Rows)
                            {
                                int serviceId = Convert.ToInt32(row.Cells["ServiceId"].Value);
                                using (MySqlCommand cmd = new MySqlCommand(insertService, con, transaction))
                                {
                                    cmd.Parameters.AddWithValue("@orderId", orderId);
                                    cmd.Parameters.AddWithValue("@serviceId", serviceId);
                                    cmd.ExecuteNonQuery();
                                }
                            }

                            // Обновление статуса расписания
                            string updateSchedule = "UPDATE Schedule SET Status = 2 WHERE idSchedule = @scheduleId";
                            using (MySqlCommand cmd = new MySqlCommand(updateSchedule, con, transaction))
                            {
                                cmd.Parameters.AddWithValue("@scheduleId", selectedScheduleId);
                                cmd.ExecuteNonQuery();
                            }

                            transaction.Commit(); // Подтверждаем транзакцию
                            MessageBox.Show("Талон успешно оформлен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // Возможность печати талона в Word
                            DialogResult printResult = MessageBox.Show(
                                "Хотите распечатать талон?",
                                "Печать талона",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (printResult == DialogResult.Yes)
                            {
                                if (selectedPatientId == 0 || selectedScheduleId == 0)
                                {
                                    MessageBox.Show("Невозможно распечатать. Выберите пациента и расписание.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }

                                try
                                {
                                    string templatePath = Path.Combine(Application.StartupPath, "talontemp.docx");

                                    if (!File.Exists(templatePath))
                                    {
                                        MessageBox.Show("Не найден шаблон талона!");
                                        return;
                                    }

                                    // Создаем папку для талонов
                                    string outputDir = Path.Combine(
                                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                        "Талоны");

                                    System.IO.Directory.CreateDirectory(outputDir);

                                    // Имя нового файла
                                    string outputFile = Path.Combine(
                                        outputDir,
                                        $"Талон_{orderId}_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

                                    // Данные
                                    string number = orderId.ToString();

                                    string fioPatient = "";
                                    string birthday = "";
                                    string chils = "";
                                    string fioDoctor = "";
                                    string speciality = "";
                                    string date = "";
                                    string time = "";

                                    using (MySqlConnection con2 = new MySqlConnection(connectionString))
                                    {
                                        con2.Open();

                                        string query = @"
                                                SELECT 
                                                    CONCAT(p.surname, ' ', p.name, ' ', p.lastname) AS patient,
                                                    DATE_FORMAT(p.date_birth, '%d.%m.%Y') AS birthday,
                                                    p.number_policy AS chils,
                                                    CONCAT(d.surname, ' ', d.name, ' ', d.lastname) AS doctor,
                                                    sp.SpecialityName AS speciality,
                                                    DATE_FORMAT(sc.Date, '%d.%m.%Y') AS date,
                                                    sc.Time AS time
                                                FROM `Order` o
                                                JOIN Patients p ON o.Patients_idPatients = p.idPatients
                                                JOIN Schedule sc ON o.schedule = sc.idSchedule
                                                JOIN Doctors d ON sc.idDoctor = d.idDoctors
                                                JOIN Speciality sp ON d.Speciality = sp.idSpeciality
                                                WHERE o.idOrder = @orderId";

                                        using (MySqlCommand cmd = new MySqlCommand(query, con2))
                                        {
                                            cmd.Parameters.AddWithValue("@orderId", orderId);

                                            using (MySqlDataReader reader = cmd.ExecuteReader())
                                            {
                                                if (reader.Read())
                                                {
                                                    fioPatient = reader["patient"].ToString();
                                                    birthday = reader["birthday"].ToString();
                                                    chils = reader["chils"].ToString();
                                                    fioDoctor = reader["doctor"].ToString();
                                                    speciality = reader["speciality"].ToString();
                                                    date = reader["date"].ToString();

                                                    if (reader["time"] != DBNull.Value)
                                                    {
                                                        time = TimeSpan.Parse(reader["time"].ToString())
                                                            .ToString(@"hh\:mm");
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    Word.Application wordApp = new Word.Application();

                                    // Открываем шаблон только для чтения
                                    Word.Document doc = wordApp.Documents.Open(
                                        templatePath,
                                        ReadOnly: true);

                                    // Сохраняем копию
                                    doc.SaveAs(outputFile);

                                    Word.Range range = doc.Content;
                                    object replaceAll = Word.WdReplace.wdReplaceAll;

                                    void Replace(string find, string value)
                                    {
                                        range.Find.Execute(
                                            FindText: find,
                                            ReplaceWith: value ?? "",
                                            Replace: replaceAll);
                                    }

                                    Replace("{number}", number);
                                    Replace("{date}", date);
                                    Replace("{time}", time);

                                    Replace("{FIOP}", fioPatient);
                                    Replace("{birthday}", birthday);
                                    Replace("{chils}", chils);

                                    Replace("{FIO}", fioDoctor);
                                    Replace("{speciality}", speciality);

                                    // Сохраняем уже готовый талон
                                    doc.Save();

                                    // Показываем пользователю готовый файл
                                    wordApp.Visible = true;

                                    MessageBox.Show(
                                        $"Талон сформирован.\nФайл сохранён:\n{outputFile}",
                                        "Успех",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(
                                        "Ошибка Word: " + ex.Message,
                                        "Ошибка",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                                }
                            }

                            // Очистка формы после сохранения
                            dataGridView2.Rows.Clear();
                            label6.Text = "Итого: 0 руб.";
                            label5.Text = "Скидка: 0 руб.";
                            label10.Text = "К оплате: 0 руб.";
                            label1.Text = "Врач: ";
                            label7.Text = "Дата приема: ";
                            label8.Text = "Время приема: ";
                            label9.Text = "Пациент: ";

                            selectedPatientId = selectedScheduleId = 0; // сброс выбора
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback(); // откат транзакции при ошибке
                            MessageBox.Show("Ошибка при сохранении: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при оформлении талона: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Удаление услуги из талона
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.Rows.Count == 0)
                {
                    MessageBox.Show("Нет услуг для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (dataGridView2.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите услугу для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Подтверждение удаления
                DialogResult result = MessageBox.Show(
                    "Вы действительно хотите удалить выбранную услугу из талона?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in dataGridView2.SelectedRows)
                    {
                        dataGridView2.Rows.Remove(row); // удаляем выбранные строки
                    }

                    UpdateTotal(); // пересчёт суммы
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении услуги: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            AddSelectedService();
        }

        // Общая логика добавления выбранной услуги в талон
        private void AddSelectedService()
        {
            try
            {
                // Проверка, выбрана ли услуга
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите услугу!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
                string serviceName = selectedRow.Cells["ServiceName"].Value.ToString();
                string price = selectedRow.Cells["Price"].Value.ToString();
                int serviceId = Convert.ToInt32(selectedRow.Cells["ServiceId"].Value);

                // Проверка на дубликат
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (Convert.ToInt32(row.Cells["ServiceId"].Value) == serviceId)
                    {
                        MessageBox.Show("Эта услуга уже добавлена в талон!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                // Добавляем услугу в список выбранных услуг
                dataGridView2.Rows.Add(serviceName, price, serviceId);
                UpdateTotal(); // пересчёт суммы
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении услуги: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
