using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class EditDoctor : Form
    {
        string connectionString;             // Строка подключения к БД
        int selectedDoctorId;                // ID редактируемого врача
        string oldSurname;                   // Старое значение фамилии
        string oldName;                      // Старое значение имени
        string oldLastname;                  // Старое значение отчества
        string oldPhone;                     // Старый телефон
        int oldSpecialityId;                 // Старый ID специальности
        private Image doctorPhoto;           // Фото врача
        private string oldPhotoFileName;     // Старое имя файла фото
        string selectedPhotoFileName;        // Выбранный файл фото
        string photoFolder;                  // Папка для фото
        string selectedPhotoFullPath;        // Полный путь выбранного фото
        bool photoDeleted = false;           // Флаг удаления фото
        string placeholderPath = Path.Combine(Application.StartupPath, "photo", "upload.png"); // Фото-заглушка
        long maxSize = 1 * 1024 * 1024; // 1 МБ

        public EditDoctor(int doctorId, Image photo)
        {
            InitializeComponent();           // Инициализация компонентов
            selectedDoctorId = doctorId;     // Сохраняем ID врача
            doctorPhoto = photo;             // Сохраняем текущее фото
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.Russian_Hyphen(sender, e); // Ограничение ввода: русские буквы и дефис
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.OnlyRussian(sender, e);    // Только русские буквы
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // Проверка, что обязательные поля заполнены
            if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                string.IsNullOrWhiteSpace(textBox2.Text) ||
                !maskedTextBox1.MaskFull ||
                comboBox2.SelectedValue == null)
            {
                MessageBox.Show("Поля не должны быть пустыми!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string newSurname = textBox1.Text.Trim();        // Новая фамилия
            string newName = textBox2.Text.Trim();           // Новое имя
            string newLastname = textBox3.Text.Trim();       // Новое отчество
            string newPhone = maskedTextBox1.Text.Trim();    // Новый телефон
            int newSpecialityId = Convert.ToInt32(comboBox2.SelectedValue); // Новая специальность

            bool photoChanged = selectedPhotoFileName != null || photoDeleted; // Проверка изменения фото

            // Проверка любых изменений
            bool changed =
                newSurname != oldSurname ||
                newName != oldName ||
                newLastname != oldLastname ||
                newPhone != oldPhone ||
                newSpecialityId != oldSpecialityId ||
                photoChanged;

            if (!changed)
            {
                MessageBox.Show("Вы не внесли никаких изменений!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string photoFileToSave = oldPhotoFileName; // Имя фото, которое будет сохранено

            if (!string.IsNullOrEmpty(selectedPhotoFileName) && !string.IsNullOrEmpty(selectedPhotoFullPath)) // Если выбрали новое фото
            {
                try
                {
                    if (!System.IO.Directory.Exists(photoFolder))
                        System.IO.Directory.CreateDirectory(photoFolder); // Создаем папку, если нет

                    string destPath = Path.Combine(photoFolder, selectedPhotoFileName);

                    bool isSameFolder =
                        string.Equals(Path.GetDirectoryName(selectedPhotoFullPath)?.TrimEnd('\\'),
                                      photoFolder.TrimEnd('\\'),
                                      StringComparison.OrdinalIgnoreCase);

                    if (isSameFolder) // Если файл уже в нужной папке
                    {
                        photoFileToSave = selectedPhotoFileName;
                    }
                    else
                    {
                        // Генерация уникального имени, если файл с таким именем существует
                        if (File.Exists(destPath))
                        {
                            string name = Path.GetFileNameWithoutExtension(selectedPhotoFileName);
                            string ext = Path.GetExtension(selectedPhotoFileName);
                            string uniqueName = $"{name}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}";
                            destPath = Path.Combine(photoFolder, uniqueName);
                            selectedPhotoFileName = uniqueName;
                        }

                        using (Image img = Image.FromFile(selectedPhotoFullPath))
                        {
                            Image finalImage = img;

                            FileInfo fileInfo = new FileInfo(selectedPhotoFullPath);

                            string fileName =
                                Path.GetFileNameWithoutExtension(selectedPhotoFileName) + ".jpg";

                            destPath = Path.Combine(photoFolder, fileName);

                            if (fileInfo.Length > maxSize)
                            {
                                finalImage = CompressImage(img, 60L);
                            }

                            if (File.Exists(destPath))
                            {
                                string nameOnly =
                                    Path.GetFileNameWithoutExtension(fileName);

                                string uniqueName =
                                    $"{nameOnly}_{DateTime.Now:yyyyMMdd_HHmmss}.jpg";

                                fileName = uniqueName;
                                destPath = Path.Combine(photoFolder, uniqueName);
                            }

                            finalImage.Save(destPath, ImageFormat.Jpeg);

                            photoFileToSave = fileName;

                            if (finalImage != img)
                                finalImage.Dispose();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении фото: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }           
            else if (photoDeleted) // Если фото удалено
            {
                if (!string.IsNullOrEmpty(oldPhotoFileName))
                {
                    string fullPath = Path.Combine(photoFolder, oldPhotoFileName);
                    if (File.Exists(fullPath))
                    {
                        try { File.Delete(fullPath); } catch { } // Удаляем старый файл
                    }
                }
                photoFileToSave = null; // Фото не сохраняем
            }

            try
            {
                // Обновление записи в БД
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    string checkQuery = @"
            SELECT COUNT(*) FROM Doctors 
            WHERE Surname = @surname 
              AND Name = @name 
              AND Lastname = @lastname 
              AND Phone_number = @phone
              AND idDoctors <> @id;";

                    using (MySqlCommand checkCmd = new MySqlCommand(checkQuery, con))
                    {
                        checkCmd.Parameters.AddWithValue("@surname", newSurname);
                        checkCmd.Parameters.AddWithValue("@name", newName);
                        checkCmd.Parameters.AddWithValue("@lastname", newLastname);
                        checkCmd.Parameters.AddWithValue("@phone", newPhone);
                        checkCmd.Parameters.AddWithValue("@id", selectedDoctorId);

                        if (Convert.ToInt32(checkCmd.ExecuteScalar()) > 0)
                        {
                            MessageBox.Show("Такая запись уже существует!", "Дубликат", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    string checkQuery2 = @"
                                            SELECT COUNT(*)
                                            FROM Doctors
                                            WHERE Phone_number = @phone
                                              AND idDoctors <> @id";
                    // Запрос проверки дубликата телефона
                    using (MySqlCommand checkCmd = new MySqlCommand(checkQuery2, con))
                    {
                        checkCmd.Parameters.AddWithValue("@phone", newPhone);
                        checkCmd.Parameters.AddWithValue("@id", selectedDoctorId);

                        int exists = Convert.ToInt32(checkCmd.ExecuteScalar());

                        if (exists > 0)
                        {
                            MessageBox.Show("Такой номер уже существует!", "Дубликат",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }

                    string query = @"
                                UPDATE Doctors
                                SET Surname = @surname,
                                    Name = @name,
                                    Lastname = @lastname,
                                    Phone_number = @phone,
                                    Speciality = @speciality,
                                    Photo = @photo
                                WHERE idDoctors = @id;";

                    using (MySqlCommand cmd = new MySqlCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@surname", newSurname);
                        cmd.Parameters.AddWithValue("@name", newName);
                        cmd.Parameters.AddWithValue("@lastname", newLastname);
                        cmd.Parameters.AddWithValue("@phone", newPhone);
                        cmd.Parameters.AddWithValue("@speciality", newSpecialityId);
                        cmd.Parameters.AddWithValue("@id", selectedDoctorId);
                        cmd.Parameters.AddWithValue("@photo", string.IsNullOrEmpty(photoFileToSave)
                            ? (object)DBNull.Value
                            : photoFileToSave);

                        cmd.ExecuteNonQuery(); // Выполняем обновление
                                               // Если фото было заменено
                        if (!string.IsNullOrEmpty(oldPhotoFileName) &&
                            oldPhotoFileName != photoFileToSave)
                        {
                            string checkPhotoQuery =
                                "SELECT COUNT(*) FROM Doctors WHERE Photo = @photo";

                            using (MySqlCommand photoCheckCmd = new MySqlCommand(checkPhotoQuery, con))
                            {
                                photoCheckCmd.Parameters.AddWithValue("@photo", oldPhotoFileName);

                                int photoUsageCount =
                                    Convert.ToInt32(photoCheckCmd.ExecuteScalar());

                                // После UPDATE текущий врач уже не ссылается на старый файл
                                if (photoUsageCount == 0)
                                {
                                    string oldPhotoPath =
                                        Path.Combine(photoFolder, oldPhotoFileName);

                                    if (File.Exists(oldPhotoPath))
                                    {
                                        try
                                        {
                                            File.Delete(oldPhotoPath);
                                        }
                                        catch
                                        {
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                MessageBox.Show("Запись успешно обновлена!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close(); // Закрываем форму
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении записи: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EditDoctor_Load(object sender, EventArgs e)
        {
            Connect connect = new Connect();
            connectionString = connect.ConnectDB(); // Получаем строку подключения
            LoadSpecialities();                      // Загружаем специальности
            LoadDoctorData();                        // Загружаем данные врача

            photoFolder = Path.Combine(Application.StartupPath, "photo");

            if (!string.IsNullOrEmpty(oldPhotoFileName))
            {
                pictureBox1.Image = new Bitmap(doctorPhoto);
            }
            else
            {
                if (File.Exists(placeholderPath))
                {
                    pictureBox1.Image = Image.FromFile(placeholderPath);
                }
            }
            label6.Text = $"Фото: {(string.IsNullOrEmpty(oldPhotoFileName) ? "нет" : oldPhotoFileName)}"; // Название фото
        }

        private void LoadSpecialities()
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string query = "SELECT idSpeciality, SpecialityName FROM Speciality;";
                    MySqlDataAdapter da = new MySqlDataAdapter(query, con);
                    DataTable t = new DataTable();
                    da.Fill(t);

                    comboBox2.DisplayMember = "SpecialityName"; // Отображаемое имя
                    comboBox2.ValueMember = "idSpeciality";     // Значение
                    comboBox2.DataSource = t;
                    comboBox2.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке специальностей: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadDoctorData()
        {
            try
            {
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    string query = @"
                    SELECT d.Surname, d.Name, d.Lastname, d.Phone_number, s.idSpeciality, d.Photo
                    FROM Doctors d
                    JOIN Speciality s ON d.Speciality = s.idSpeciality
                    WHERE d.idDoctors = @id;";
                    MySqlCommand cmd = new MySqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@id", selectedDoctorId);

                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            oldSurname = reader["Surname"].ToString();      // Сохраняем старые значения
                            oldName = reader["Name"].ToString();
                            oldLastname = reader["Lastname"].ToString();
                            oldPhone = reader["Phone_number"].ToString();
                            oldSpecialityId = Convert.ToInt32(reader["idSpeciality"]);
                            oldPhotoFileName = reader["Photo"] == DBNull.Value ? "" : reader["Photo"].ToString();

                            textBox1.Text = oldSurname;     // Заполняем текстбоксы
                            textBox2.Text = oldName;
                            textBox3.Text = oldLastname;
                            maskedTextBox1.Text = oldPhone;
                            comboBox2.SelectedValue = oldSpecialityId;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных врача: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void UploadPhoto()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Изображения (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png";
                ofd.Title = "Выберите новое фото";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(ofd.FileName);

                        using (Image originalImage = Image.FromFile(ofd.FileName))
                        {
                            // Проверка 400x400
                            if (originalImage.Width != 400 || originalImage.Height != 400)
                            {
                                MessageBox.Show("Изображение должно быть 400x400 пикселей!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            Image finalImage;

                            // если больше 1MB — сжимаем
                            if (fileInfo.Length > maxSize)
                            {
                                double sizeMb = fileInfo.Length / 1024.0 / 1024.0;

                                DialogResult result = MessageBox.Show(
                                    $"Размер изображения составляет {sizeMb:F2} МБ.\n" +
                                    $"Допустимый размер — не более 1 МБ.\n\n" +
                                    $"Сжать изображение?",
                                    "Превышен размер файла",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question);

                                if (result == DialogResult.No)
                                    return;

                                finalImage = CompressImage(originalImage, 60L);

                                // определяем размер после сжатия
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    finalImage.Save(ms, ImageFormat.Jpeg);

                                    double compressedSizeMb = ms.Length / 1024.0 / 1024.0;

                                    MessageBox.Show(
                                        $"Изображение успешно сжато.\n" +
                                        $"Новый размер: {compressedSizeMb:F2} МБ.",
                                        "Сжатие выполнено",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);
                                }
                            }
                            else
                            {
                                finalImage = new Bitmap(originalImage);
                            }

                            pictureBox1.Image?.Dispose();
                            pictureBox1.Image = finalImage;
                        }

                        selectedPhotoFullPath = ofd.FileName;
                        selectedPhotoFileName = Path.GetFileName(ofd.FileName);
                        photoDeleted = false; // пользователь выбрал новое фото
                        label6.Text = $"Фото: {selectedPhotoFileName}";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при загрузке изображения: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UploadPhoto();
        }

        private Image CompressImage(Image image, long quality = 75L)
        {
            ImageCodecInfo jpgEncoder = ImageCodecInfo.GetImageDecoders()
                .First(c => c.FormatID == ImageFormat.Jpeg.Guid);

            EncoderParameters encoderParameters = new EncoderParameters(1);
            encoderParameters.Param[0] = new EncoderParameter(Encoder.Quality, quality);

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, jpgEncoder, encoderParameters);
                return Image.FromStream(new MemoryStream(ms.ToArray()));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(oldPhotoFileName))
            {
                MessageBox.Show("У врача нет сохранённого фото.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show("Удалить фото?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            Image oldImage = pictureBox1.Image;
            pictureBox1.Image = Image.FromFile(placeholderPath);
            oldImage?.Dispose();
            label6.Text = "Фото: нет";
            selectedPhotoFileName = null;
            selectedPhotoFullPath = null;
            photoDeleted = true; // Отмечаем, что фото будет удалено

            MessageBox.Show("Фото будет удалено после сохранения.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
