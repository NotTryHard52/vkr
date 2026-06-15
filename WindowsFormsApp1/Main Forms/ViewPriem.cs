using MySql.Data.MySqlClient; // Подключение MySQL клиента
using System;
using System.Data;
using System.IO; // Работа с файлами
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word; // Используем Word Interop для печати чека

namespace WindowsFormsApp1
{
    public partial class ViewPriem : Form
    {
        private int orderId; // ID текущего приёма
        string connectionString; // Строка подключения к БД
        decimal totalSum = 0; // Сумма всех услуг
        private bool isGlav; // Флаг, если пользователь главный (ограничения на редактирование)

        // Конструктор формы
        public ViewPriem(int orderId, bool isGlav = false)
        {
            InitializeComponent();
            this.orderId = orderId;
            this.isGlav = isGlav;
        }

        // Событие загрузки формы
        private void ViewPriem_Load(object sender, EventArgs e)
        {
            try
            {
                Connect connect = new Connect();
                connectionString = connect.ConnectDB(); // Получаем строку подключения
                LoadOrderData(); // Загружаем данные о приёме
                LoadServices();  // Загружаем услуги в таблицу
                LoadStatuses();  // Загружаем статусы в comboBox
                CalculateTotal(); // Вычисляем итоговую сумму

                if (isGlav) { DisableForGlav(); } // Блокировка для главного пользователя

                // Блокировка элементов при определённых статусах
                if (comboBox1.Text == "Завершен" || comboBox1.Text == "Отменен")
                {
                    DisableEditing();
                }
                if (comboBox1.Text == "Создан" || comboBox1.Text == "Отменен")
                {
                    button5.Enabled = false; // Печать чека недоступна
                }

                var hoverEffect = new HoverDataGridView(dataGridView1); // Эффект при наведении на DataGridView
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Блокировка элементов для главного пользователя
        private void DisableForGlav()
        {
            button1.Visible = false; // Добавить услугу
            button2.Visible = false; // Удалить услугу
            button4.Visible = false; // Завершить приём
            button6.Visible = false; // Отменить приём
            button5.Visible = false; // Печать чека
            comboBox1.Enabled = false;
            dataGridView1.Enabled = false;
        }

        // Полная блокировка редактирования
        private void DisableEditing()
        {
            dataGridView1.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            button4.Enabled = false;
            button6.Enabled = false;
            comboBox1.Enabled = false;
        }

        // Загрузка данных о приёме из БД
        private void LoadOrderData()
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string query = @"
                        SELECT 
                            o.idOrder,
                            o.sum,
                            o.Discount,
                            o.TotalSum,
                            CONCAT(p.surname, ' ', p.name, ' ', p.lastname) AS patient_name,
                            CONCAT(d.surname, ' ', d.name, ' ', d.lastname) AS doctor_name,
                            DATE_FORMAT(sc.date, '%d.%m.%Y') AS date,
                            DATE_FORMAT(sc.time, '%H:%i') AS time,
                            st.name AS status
                        FROM `Order` o
                        JOIN Schedule sc ON o.Schedule = sc.idSchedule
                        JOIN Doctors d ON sc.idDoctor = d.idDoctors
                        JOIN Patients p ON o.Patients_idPatients = p.idPatients
                        JOIN StatusesPriem st ON o.Status = st.idStatusesPriem
                        WHERE o.idOrder = @orderId;";

                    using (MySqlCommand cmd = new MySqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@orderId", orderId);
                        using (MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // Заполнение меток формы данными из БД
                                label_number.Text = "Номер талона: " + reader["idOrder"].ToString();
                                label_patient.Text = "Пациент: " + reader["patient_name"].ToString();
                                label_doctor.Text = "Врач: " + reader["doctor_name"].ToString();
                                label_data.Text = "Дата: " + reader["date"].ToString();
                                label_time.Text = "Время: " + reader["time"].ToString();
                                comboBox1.Text = reader["status"].ToString();
                                // Сумма, скидка и итог к оплате
                                decimal sum = reader.GetDecimal("sum");
                                decimal discount = reader.GetDecimal("Discount");
                                decimal total = reader.GetDecimal("TotalSum");

                                label_total.Text = $"Итого: {total:N2} руб.";
                                label5.Text = $"Скидка: {discount:N2} руб.";
                                label10.Text = $"К оплате: {total:N2} руб.";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных о приёме: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Загрузка услуг из БД в DataGridView
        private void LoadServices()
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string serviceQuery = @"
                    SELECT s.Name AS 'Услуга', s.Price AS 'Цена'
                    FROM OrderServices os
                    INNER JOIN Services s ON os.ServicesId = s.idServices
                    WHERE os.OrderId = @orderId;";
                    using (MySqlCommand cmd = new MySqlCommand(serviceQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@orderId", orderId);
                        DataTable dt = new DataTable();
                        new MySqlDataAdapter(cmd).Fill(dt);
                        dataGridView1.DataSource = dt;
                        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect; // Выбор полной строки
                        dataGridView1.Columns["Услуга"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        dataGridView1.Columns["Цена"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке услуг: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Загрузка всех статусов приёма в comboBox
        private void LoadStatuses()
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string statusQuery = "SELECT name FROM StatusesPriem;";
                    using (MySqlCommand cmd = new MySqlCommand(statusQuery, con))
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        comboBox1.Items.Clear();
                        while (reader.Read())
                            comboBox1.Items.Add(reader["name"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке статусов: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Вычисление общей суммы услуг
        private void CalculateTotal()
        {
            try
            {
                totalSum = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells["Цена"]?.Value != null && decimal.TryParse(row.Cells["Цена"].Value.ToString(), out decimal value))
                        totalSum += value;
                }
                label_total.Text = $"Итого: {totalSum:F2} руб.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при вычислении суммы: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // Завершение приёма
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Нельзя закрыть приём без оказанных услуг!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string statusQuery = "SELECT idStatusesPriem FROM StatusesPriem WHERE name='Завершен' LIMIT 1;";
                    int statusId = Convert.ToInt32(new MySqlCommand(statusQuery, con).ExecuteScalar());

                    string updateQuery = "UPDATE `Order` SET Status=@status WHERE idOrder=@orderId;";
                    using (MySqlCommand cmd = new MySqlCommand(updateQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@status", statusId);
                        cmd.Parameters.AddWithValue("@orderId", orderId);
                        cmd.ExecuteNonQuery();
                    }

                    comboBox1.Text = "Завершен";

                    // Блокировка элементов формы
                    dataGridView1.Enabled = false;
                    button1.Enabled = false;
                    button2.Enabled = false;
                    button4.Enabled = false;
                    button5.Enabled = true; // Разрешена печать чека
                    button6.Enabled = false;

                    MessageBox.Show("Приём завершён!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при завершении приёма: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Кнопка "Добавить услугу"
        private void button1_Click(object sender, EventArgs e)
        {
            AddServiceToOrder(orderId);
            LoadServices();
            CalculateTotal();
        }

        // Кнопка "Удалить услугу"
        private void button2_Click(object sender, EventArgs e)
        {
            RemoveSelectedService(orderId);
            LoadServices();
            CalculateTotal();
        }

        // Метод добавления услуги к приёму
        private void AddServiceToOrder(int orderId)
        {
            try
            {
                Services servicesForm = new Services(true);
                if (servicesForm.ShowDialog() == DialogResult.OK)
                {
                    int serviceId = servicesForm.SelectedServiceId;
                    string serviceName = servicesForm.SelectedServiceName;
                    decimal servicePrice = servicesForm.SelectedServicePrice;

                    DataTable dt = (DataTable)dataGridView1.DataSource;

                    // Проверка на дублирование услуги
                    foreach (DataRow existingRow in dt.Rows)
                    {
                        if (existingRow["Услуга"].ToString() == serviceName)
                        {
                            MessageBox.Show("Эта услуга уже добавлена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    // Добавление услуги в таблицу
                    DataRow newRow = dt.NewRow();
                    newRow["Услуга"] = serviceName;
                    newRow["Цена"] = servicePrice;
                    dt.Rows.Add(newRow);

                    // Добавление услуги в базу
                    using (MySqlConnection con = new MySqlConnection(connectionString))
                    {
                        con.Open();
                        string insertQuery = "INSERT INTO OrderServices (OrderId, ServicesId) VALUES (@orderId, @serviceId)";
                        using (MySqlCommand cmd = new MySqlCommand(insertQuery, con))
                        {
                            cmd.Parameters.AddWithValue("@orderId", orderId);
                            cmd.Parameters.AddWithValue("@serviceId", serviceId);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    CalculateTotal();
                    UpdateOrderTotalsInDatabase(); // Обновляем общие суммы в БД
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении услуги: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод удаления выбранной услуги
        private void RemoveSelectedService(int orderId)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите услугу для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    "Вы действительно хотите удалить выбранную услугу из талона?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string serviceName = dataGridView1.SelectedRows[0].Cells["Услуга"].Value.ToString();

                    using (MySqlConnection con = new MySqlConnection(connectionString))
                    {
                        con.Open();
                        string deleteQuery = @"
                        DELETE FROM OrderServices 
                        WHERE OrderId = @orderId 
                          AND ServicesId = (SELECT idServices FROM Services WHERE Name = @serviceName LIMIT 1)
                        LIMIT 1;";

                        using (MySqlCommand cmd = new MySqlCommand(deleteQuery, con))
                        {
                            cmd.Parameters.AddWithValue("@orderId", orderId);
                            cmd.Parameters.AddWithValue("@serviceName", serviceName);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    LoadServices();
                    CalculateTotal();
                    UpdateOrderTotalsInDatabase();

                    MessageBox.Show("Услуга удалена.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении услуги: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Закрытие формы
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Обновление итоговой суммы, скидки и оплаты в БД
        private void UpdateOrderTotalsInDatabase()
        {
            try
            {
                decimal sum = 0;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells["Цена"]?.Value != null && decimal.TryParse(row.Cells["Цена"].Value.ToString(), out decimal value))
                        sum += value;
                }

                decimal discount = 0;
                if (sum >= 5000) discount = sum * 0.05m; // Скидка 5% при сумме >5000
                decimal total = sum - discount;

                // Обновление меток
                label_total.Text = $"Итого: {total:N2} руб.";
                label5.Text = $"Скидка: {discount:N2} руб.";
                label10.Text = $"К оплате: {total:N2} руб.";

                // Обновление БД
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string updateQuery = @"
                    UPDATE `Order` 
                    SET sum=@sum, Discount=@discount, TotalSum=@total 
                    WHERE idOrder=@orderId";
                    using (MySqlCommand cmd = new MySqlCommand(updateQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@sum", sum);
                        cmd.Parameters.AddWithValue("@discount", discount);
                        cmd.Parameters.AddWithValue("@total", total);
                        cmd.Parameters.AddWithValue("@orderId", orderId);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении итогов в БД: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Печать чека через Word
        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Нет услуг для печати чека!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string templatePath = Path.Combine(Application.StartupPath, "ordertemp.docx");
                string birthday = "";
                string snils = "";
                string speciality = "";

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Не найден шаблон чека!", "Ошибка");
                    return;
                }

                Word.Application wordApp = new Word.Application();
                Word.Document doc = wordApp.Documents.Open(templatePath);
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    string query = @"
                            SELECT 
                                    DATE_FORMAT(p.date_birth, '%d.%m.%Y') AS birthday,
                                    p.number_policy AS snils,
                                    CONCAT(d.surname, ' ', d.name, ' ', d.lastname) AS doctor_name,
                                    sp.SpecialityName AS speciality
                                FROM `Order` o
                                JOIN Schedule sc ON o.Schedule = sc.idSchedule
                                JOIN Doctors d ON sc.idDoctor = d.idDoctors
                                JOIN Patients p ON o.Patients_idPatients = p.idPatients
                                JOIN Speciality sp ON d.Speciality = sp.idSpeciality
                                WHERE o.idOrder = @orderId";

                    using (MySqlCommand cmd = new MySqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@orderId", orderId);

                        using (MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                birthday = reader["birthday"].ToString();
                                snils = reader["snils"].ToString();
                                speciality = reader["speciality"].ToString();
                            }
                        }
                    }
                }

                // Данные
                string orderNumber = label_number.Text.Replace("Номер талона: ", "");
                string patientName = label_patient.Text.Replace("Пациент: ", "");
                string doctorName = label_doctor.Text.Replace("Врач: ", "");
                string date = label_data.Text.Replace("Дата: ", "");
                string time = label_time.Text.Replace("Время: ", "");

                string total = label_total.Text.Replace("Итого: ", "");
                string discount = label5.Text.Replace("Скидка: ", "");
                string final = label10.Text.Replace("К оплате: ", "");

                Word.Range range = doc.Content;

                object replaceAll = Word.WdReplace.wdReplaceAll;

                // Замена плейсхолдеров
                range.Find.Execute("{number}", ReplaceWith: orderNumber, Replace: replaceAll);
                range.Find.Execute("{FIOP}", ReplaceWith: patientName, Replace: replaceAll);
                range.Find.Execute("{FIO}", ReplaceWith: doctorName, Replace: replaceAll);
                range.Find.Execute("{orderdate}", ReplaceWith: date, Replace: replaceAll);
                range.Find.Execute("{birthday}", ReplaceWith: birthday, Replace: replaceAll);
                range.Find.Execute("{chils}", ReplaceWith: snils, Replace: replaceAll);
                range.Find.Execute("{speciality}", ReplaceWith: speciality, Replace: replaceAll);

                range.Find.Execute("{total}", ReplaceWith: total, Replace: replaceAll);
                range.Find.Execute("{discount}", ReplaceWith: discount, Replace: replaceAll);
                range.Find.Execute("{pay}", ReplaceWith: final, Replace: replaceAll);

                // ===== ФОРМИРУЕМ СПИСОК УСЛУГ =====
                StringBuilder servicesBuilder = new StringBuilder();
                StringBuilder costBuilder = new StringBuilder();

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    servicesBuilder.Append(
                        dataGridView1.Rows[i].Cells["Услуга"].Value.ToString() + "\v");

                    costBuilder.Append(
                        dataGridView1.Rows[i].Cells["Цена"].Value.ToString() + " руб." + "\v");
                }

                // Замена плейсхолдеров
                range.Find.Execute("{nameservices}",
                    ReplaceWith: servicesBuilder.ToString(),
                    Replace: replaceAll);

                range.Find.Execute("{cost}",
                    ReplaceWith: costBuilder.ToString(),
                    Replace: replaceAll);

                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // имя файла
                string pdfPath = Path.Combine(
                    documentsPath,
                    $"Чек_{orderNumber}.pdf");

                // экспорт в PDF
                doc.ExportAsFixedFormat(
                    pdfPath,
                    Word.WdExportFormat.wdExportFormatPDF);

                // закрываем Word
                doc.Close(false);
                wordApp.Quit();

                // открываем PDF
                System.Diagnostics.Process.Start(pdfPath);

                MessageBox.Show(
                    "PDF-чек успешно создан!",
                    "Успех",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        // Отмена приёма
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    // Получаем id статуса "Отменен" для Order
                    int orderStatusId = Convert.ToInt32(new MySqlCommand(
                        "SELECT idStatusesPriem FROM StatusesPriem WHERE Name='Отменен' LIMIT 1;", con).ExecuteScalar());

                    // Получаем ScheduleId для текущего Order
                    MySqlCommand scheduleCmd = new MySqlCommand(
                        "SELECT Schedule FROM `Order` WHERE idOrder=@orderId LIMIT 1;", con);
                    scheduleCmd.Parameters.AddWithValue("@orderId", orderId);
                    int scheduleId = Convert.ToInt32(scheduleCmd.ExecuteScalar());

                    // Получаем id статуса "Свободно" для Schedule
                    int freeScheduleStatusId = Convert.ToInt32(new MySqlCommand(
                        "SELECT idStatuses FROM Statuses WHERE StatusName='Свободно' LIMIT 1;", con).ExecuteScalar());

                    if (orderStatusId == 0 || scheduleId == 0 || freeScheduleStatusId == 0)
                    {
                        MessageBox.Show("Ошибка: не найден нужный статус или расписание.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Обновление Schedule — перевод в статус "Свободно"
                    MySqlCommand updateScheduleCmd = new MySqlCommand(
                        "UPDATE Schedule SET Status=@scheduleStatusId WHERE idSchedule=@scheduleId;", con);
                    updateScheduleCmd.Parameters.AddWithValue("@scheduleStatusId", freeScheduleStatusId);
                    updateScheduleCmd.Parameters.AddWithValue("@scheduleId", scheduleId);
                    updateScheduleCmd.ExecuteNonQuery();

                    // Обновление Order — перевод в статус "Отменен"
                    MySqlCommand updateOrderCmd = new MySqlCommand(
                        "UPDATE `Order` SET Status=@orderStatusId WHERE idOrder=@orderId;", con);
                    updateOrderCmd.Parameters.AddWithValue("@orderStatusId", orderStatusId);
                    updateOrderCmd.Parameters.AddWithValue("@orderId", orderId);
                    updateOrderCmd.ExecuteNonQuery();

                    comboBox1.Text = "Отменен";

                    // Блокировка элементов формы
                    dataGridView1.Enabled = false;
                    button1.Enabled = false;
                    button2.Enabled = false;
                    button4.Enabled = false;
                    button6.Enabled = false;

                    MessageBox.Show("Приём отменен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при отмене приёма: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
