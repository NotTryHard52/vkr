using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Doctor : Form
    {
        string connectionString;               // Строка подключения к БД
        DataTable doctorsTable;                // Таблица для хранения данных врачей
        int selectedId = -1;                   // ID выбранного врача (-1 = не выбран)


        public Doctor()
        {
            InitializeComponent();             // Инициализация компонентов формы
            this.Resize += Doctor_Resize;
        }

        private void Doctor_Resize(object sender, EventArgs e)
        {
            UpdateCardLayout();
        }

        private void UpdateCardLayout()
        {
            foreach (Control ctrl in flowLayoutPanel1.Controls)
            {
                ctrl.Width = 548;
                ctrl.Height = 250;
            }
        }

        private void Doctor_Load(object sender, EventArgs e)
        {
            try
            {
                typeof(FlowLayoutPanel)
                        .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        ?.SetValue(flowLayoutPanel1, true, null);

                FillSpecialties();                 // Заполнение списка специальностей
                comboBox2.SelectedIndex = 0;       // Значение сортировки по умолчанию
                comboBox1.SelectedIndex = 0;       // Фильтр специальностей по умолчанию
                comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged; // Обработчик смены фильтра
                textBox5.TextChanged += textBox5_TextChanged;                    // Обработчик ввода поиска
                LoadDoctor();                      // Загрузка списка врачей
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadDoctor()
        {
            try
            {
                Connect connect = new Connect();
                connectionString = connect.ConnectDB(); // Получаем строку подключения

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();                         // Открытие соединения

                    doctorsTable = new DataTable();     // Таблица для загрузки данных

                    // SQL-запрос — выбираем врачей + название специальности
                    MySqlCommand cmd = new MySqlCommand(
                        "SELECT d.idDoctors, d.Surname, d.Name, d.Lastname, d.Phone_number, d.Photo, s.SpecialityName " +
                        "FROM Doctors d JOIN Speciality s ON d.Speciality = s.idSpeciality;", con);

                    MySqlDataAdapter da = new MySqlDataAdapter(cmd); // Адаптер данных
                    da.Fill(doctorsTable);                           // Заполняем DataTable
                }

                LoadDoctorPhotos(doctorsTable);       // Загружаем фото врачей в таблицу

                DisplayCards(doctorsTable); // Привязка таблицы к DataGridView
                groupBox2.Text = $"Количество записей: {doctorsTable.Rows.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DisplayCards(DataTable table)
        {
            try
            {
                flowLayoutPanel1.SuspendLayout();
                flowLayoutPanel1.Controls.Clear();

                foreach (DataRow row in table.Rows)
                {
                    DoctorCard card = new DoctorCard();

                    int id = Convert.ToInt32(row["idDoctors"]);

                    string fio = $"{row["Surname"]} {row["Name"]} {row["Lastname"]}";
                    string phone = row["Phone_number"].ToString();
                    string spec = row["SpecialityName"].ToString();
                    Image photo = row["PhotoImage"] as Image;

                    card.SetData(id, fio, phone, spec, photo);
                    card.EditClicked += Card_EditClicked;
                    card.DeleteClicked += Card_DeleteClicked;
                    card.MouseClick += (s, e) => selectedId = id;

                    flowLayoutPanel1.Controls.Add(card);
                }

                flowLayoutPanel1.ResumeLayout();
                groupBox2.Text = $"Количество записей: {table.Rows.Count}";
                UpdateCardLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отображении карточек:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadDoctorPhotos(DataTable table)
        {
            if (table == null) return;              // Проверка на null

            if (!table.Columns.Contains("Фото"))    // Если нет колонки "Фото"
                table.Columns.Add("PhotoImage", typeof(Image));

            string photoFolder = Path.Combine(Application.StartupPath, "photo"); // Папка с фото

            foreach (DataRow row in table.Rows)
            {
                string fileName = "not-image.png";  // Фото по умолчанию

                if (row["Photo"] != DBNull.Value && !string.IsNullOrWhiteSpace(row["Photo"].ToString()))
                    fileName = row["Photo"].ToString();

                string photoPath = Path.Combine(photoFolder, fileName); // Полный путь

                if (!File.Exists(photoPath)) // Если фото не существует
                    photoPath = Path.Combine(photoFolder, "not-image.png");

                try
                {
                    // Загружаем изображение через FileStream, чтобы избежать блокировок
                    using (FileStream fs = new FileStream(photoPath, FileMode.Open, FileAccess.Read))
                    {
                        using (var tempImg = Image.FromStream(fs))
                        {
                            row["PhotoImage"] = new Bitmap(tempImg); // Добавляем фото в таблицу
                        }
                    }
                }
                catch
                {
                    row["PhotoImage"] = new Bitmap(100, 100); // В случае ошибки — пустая картинка
                }
            }
        }

        private void FillSpecialties()
        {
            try
            {
                Connect connect = new Connect();
                connectionString = connect.ConnectDB(); // Получаем строку подключения

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    string query = "SELECT DISTINCT SpecialityName FROM speciality;"; // Список специальностей
                    MySqlCommand cmd = new MySqlCommand(query, con);

                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        comboBox1.Items.Clear();        // Очищаем список
                        comboBox1.Items.Add("Все специальности");     // Пункт для отображения всех врачей

                        while (reader.Read())
                        {
                            string specialty = reader["SpecialityName"].ToString();
                            comboBox1.Items.Add(specialty); // Добавляем специальность
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке специальностей:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort(); // Сортировка
        }

        private void ApplyFilterAndSort()
        {
            try
            {
                if (doctorsTable == null) return;

                string filterExpr = ""; // Строка фильтра

                // Фильтр по специальности
                string selectedSpecialty =
                            comboBox1.SelectedIndex > 0
                            ? comboBox1.SelectedItem.ToString()
                            : null;
                if (!string.IsNullOrEmpty(selectedSpecialty) && selectedSpecialty != "Все специальности")
                {
                    filterExpr = $"SpecialityName = '{selectedSpecialty.Replace("'", "''")}'";
                }

                // Фильтр по фамилии
                string searchText = textBox5.Text.Trim().Replace("'", "''");
                if (!string.IsNullOrEmpty(searchText))
                {
                    if (!string.IsNullOrEmpty(filterExpr))
                        filterExpr += " AND ";

                    filterExpr += $"Surname LIKE '%{searchText}%'";
                }

                // Сортировка
                string sortExpr = "";
                if (comboBox2.SelectedIndex == 1)
                    sortExpr = "Surname ASC";
                else if (comboBox2.SelectedIndex == 2)
                    sortExpr = "Surname DESC";

                DataView dv = doctorsTable.DefaultView; // Представление таблицы
                dv.RowFilter = filterExpr;              // Применяем фильтр
                dv.Sort = string.IsNullOrWhiteSpace(sortExpr) ? null : sortExpr;                     // Применяем сортировку

                DisplayCards(dv.ToTable());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при фильтрации/сортировке:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort(); // Обновляем отображение при смене специальности
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort(); // Фильтр по фамилии
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddDoctor ad = new AddDoctor(); // Форма добавления врача
            ad.ShowDialog();                // Открываем модально

            LoadDoctor();                   // Перезагружаем таблицу
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.Russian_Hyphen(sender, e); // Ограничиваем ввод ФИО и дефисов
        }

        private void button5_Click(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = 0; // Сброс сортировки
            comboBox1.SelectedIndex = 0; // Сброс фильтра
            textBox5.Text = "";          // Очистка поиска
        }


        private void Card_EditClicked(object sender, int id)
        {
            try
            {
                Image photo = null;

                DataRow[] rows = doctorsTable.Select($"idDoctors = {id}");
                if (rows.Length > 0 && rows[0]["PhotoImage"] is Image img)
                {
                    photo = new Bitmap(img);
                }

                EditDoctor editForm = new EditDoctor(id, photo);
                editForm.ShowDialog();

                LoadDoctor();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при редактировании врача:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Card_DeleteClicked(object sender, int id)
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    // Проверка использования в расписании
                    string checkQuery = "SELECT COUNT(*) FROM Schedule WHERE idDoctor = @doctorId";
                    using (MySqlCommand checkCmd = new MySqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@doctorId", id);
                        int count = Convert.ToInt32(checkCmd.ExecuteScalar());

                        if (count > 0)
                        {
                            MessageBox.Show("Нельзя удалить врача, он используется в расписании!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    if (MessageBox.Show("Удалить врача?", "Подтверждение",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question) != DialogResult.Yes)
                        return;

                    // Получаем имя фото
                    string photoFileName = "";

                    string photoQuery = "SELECT Photo FROM Doctors WHERE idDoctors = @id";
                    using (MySqlCommand photoCmd = new MySqlCommand(photoQuery, con))
                    {
                        photoCmd.Parameters.AddWithValue("@id", id);

                        object result = photoCmd.ExecuteScalar();

                        if (result != null && result != DBNull.Value)
                        {
                            photoFileName = result.ToString();
                        }
                    }

                    // Удаляем врача из БД
                    string deleteQuery = "DELETE FROM Doctors WHERE idDoctors = @id";
                    using (MySqlCommand cmd = new MySqlCommand(deleteQuery, con))
                    {
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                    }

                    // Удаляем фото из папки
                    if (!string.IsNullOrWhiteSpace(photoFileName))
                    {
                        string photoPath = Path.Combine(Application.StartupPath, "photo", photoFileName);

                        if (File.Exists(photoPath))
                        {
                            try
                            {
                                File.Delete(photoPath);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Врач удалён, но не удалось удалить фото:\n{ex.Message}", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }

                LoadDoctor();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении врача:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
