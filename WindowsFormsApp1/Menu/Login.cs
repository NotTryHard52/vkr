using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Login : Form
    {
        // Строка подключения к базе данных
        string connectionString;
        string chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
        string Captcha_current;
        Random random = new Random();
        int loginAttempts = 0;
        bool captchaRequired = false;
        DateTime blockUntil = DateTime.MinValue;

        public Login()
        {
            InitializeComponent();
        }

        // Кнопка "Войти"
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Создаем объект для подключения
                Connect connect = new Connect();
                connectionString = connect.ConnectDB();

                // Проверка на заполнение полей логин и пароль
                if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text))
                {
                    MessageBox.Show("Заполните все поля!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Встроенный вход без БД
                if (textBox1.Text == "admin" && textBox2.Text == "admin")
                {
                    string FIO = "";
                    int userId = 0;
                    int role = 1;

                    textBox1.Clear();
                    textBox2.Clear();

                    this.Hide();

                    Form nextForm = new Menu(FIO, userId, role, true);
                    nextForm.FormClosed += (s, args) => this.Show();
                    nextForm.Show();

                    return;
                }

                if (DateTime.Now < blockUntil)
                {
                    int secondsLeft = (int)(blockUntil - DateTime.Now).TotalSeconds;
                    MessageBox.Show($"Попробуйте снова через {secondsLeft} сек.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (captchaRequired)
                {
                    if (textBox3.Text != Captcha_current)
                    {
                        MessageBox.Show("Неверная капча!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        blockUntil = DateTime.Now.AddSeconds(10);
                        Captcha_Load();
                        textBox3.Clear();

                        return;
                    }
                }

                string login = textBox1.Text;
                string password = textBox2.Text;

                // Хешируем введённый пароль с помощью SHA256
                string hash_password;
                using (SHA256 sha = SHA256.Create())
                {
                    byte[] hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
                    hash_password = BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
                }

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();

                    // Запрос на получение данных пользователя по логину
                    MySqlCommand cmd = new MySqlCommand(
                        "SELECT idUsers, Password, Role, Surname, Name, Lastname " +
                        "FROM users WHERE login = @login", con);
                    cmd.Parameters.AddWithValue("@login", login);

                    MySqlDataAdapter sda = new MySqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);

                    // Если пользователь не найден
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("Пользователь не найден!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        textBox1.Clear();
                        textBox2.Clear();
                        return;
                    }

                    DataRow userRow = dt.Rows[0];

                    int userId = Convert.ToInt32(userRow["idUsers"]);
                    string dbPasswordHash = userRow["Password"].ToString();
                    int role = Convert.ToInt32(userRow["Role"].ToString());
                    string FIO = $"{userRow["Surname"]} {userRow["Name"]} {userRow["Lastname"]}";

                    // Проверка пароля
                    if (hash_password != dbPasswordHash)
                    {
                        loginAttempts++;

                        MessageBox.Show("Неверный пароль!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        textBox1.Clear();
                        textBox2.Clear();

                        if (loginAttempts >= 1)
                        {
                            captchaRequired = true;
                            pictureBox3.Visible = true;
                            button3.Visible = true;
                            label3.Visible = true;
                            textBox3.Visible = true;
                            pictureBox3.Visible = true;
                            pictureBox3.Refresh();
                            Captcha_Load();

                            this.MinimumSize = new Size(302, 593);
                            this.Height = 593;
                            pictureBox3.Location = new Point(12, 278);
                            button3.Location = new Point(12, 341);
                            label3.Location = new Point(8, 374);
                            textBox3.Location = new Point(12, 397);
                            button1.Location = new Point(12, 429);
                        }

                        return;
                    }

                    // Успешный вход — открытие формы в зависимости от роли пользователя
                    loginAttempts = 0;
                    captchaRequired = false;
                    textBox1.Clear();
                    textBox2.Clear();
                    this.Hide();
                    Form nextForm = null;
                    if (role == 1) // Администратор
                    {
                        nextForm = new Menu(FIO, userId, role, false);
                    }
                    else if (role == 2) // Регистратор
                    {
                        nextForm = new Menu_registrator(FIO, userId, role);
                    }
                    else if (role == 3) // Главный врач
                    {
                        nextForm = new Form1(FIO, userId, role);
                    }
                    nextForm.FormClosed += (s, args) => this.Show();
                    nextForm.Show();
                    captchaRequired = false;
                    pictureBox3.Visible = false;
                    button3.Visible = false;
                    label3.Visible = false;
                    textBox3.Visible = false;
                    textBox3.Clear();
                    pictureBox3.Visible = false;
                    button1.Location = new Point(12, 278);
                    this.MinimumSize = new Size(302, 447);
                    this.Height = 447;
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка базы данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        string GenerateCaptcha(int length)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < length; i++)
            {
                sb.Append(chars[random.Next(chars.Length)]);
            }

            return sb.ToString();
        }

        private void Captcha_Load()
        {
            try
            {
                if (pictureBox3.Width <= 0 || pictureBox3.Height <= 0)
                    return;

                Captcha_current = GenerateCaptcha(5);
                int width = Math.Max(pictureBox3.Width, 150);
                int height = Math.Max(pictureBox3.Height, 50);

                Bitmap bitmap = new Bitmap(width, height);
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.White);

                    for (int i = 0; i < bitmap.Width * bitmap.Height / 3; i++)
                    {
                        int xNoise = random.Next(bitmap.Width);
                        int yNoise = random.Next(bitmap.Height);
                        bitmap.SetPixel(xNoise, yNoise, Color.FromArgb(random.Next(256), random.Next(256), random.Next(256)));
                    }
                    using (Font font = new Font("Arial", 24, FontStyle.Bold))
                    {
                        int x = 10;
                        foreach (char c in Captcha_current)
                        {
                            float offsetY = random.Next(-5, 6);
                            g.DrawString(c.ToString(), font, Brushes.Black, new PointF(x, 10 + offsetY));

                            int charWidth = (int)g.MeasureString(c.ToString(), font).Width;
                            int charHeight = (int)g.MeasureString(c.ToString(), font).Height;

                            int lineY = (int)(10 + offsetY + random.Next(charHeight / 3, (charHeight * 2) / 3));
                            Pen pen = new Pen(Color.FromArgb(random.Next(256), random.Next(256), random.Next(256)), 2);
                            g.DrawLine(pen, x, lineY, x + charWidth, lineY);

                            x += random.Next(20, 35);
                        }
                    }
                }
                pictureBox3.Image = bitmap;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка генерации капчи: " + ex.Message);
            }
        }

        // Кнопка "Выход"
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                BackupClass.CreateBackupWithDialog(AppDomain.CurrentDomain.BaseDirectory);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка дампа");
                MessageBox.Show(ex.Message);
            }
            Application.Exit();
        }

        // Клик по иконке настроек
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Settingscs sett = new Settingscs();
            this.Hide();
            sett.ShowDialog();
            this.Show();
        }

        // Загрузка формы Login
        private void Login_Load(object sender, EventArgs e)
        {
            Connect connect = new Connect();
            connectionString = connect.ConnectNoDB(); // Подключение без указания базы данных
            MySqlConnection con = new MySqlConnection(connectionString);

            try
            {
                con.Open();
                pictureBox3.Visible = false;
                button3.Visible = false;
                label3.Visible = false;
                textBox3.Visible = false;
            }
            catch (Exception)
            {
                // Ошибка подключения — предложить настройки
                DialogResult res = MessageBox.Show(
                    $"Ошибка подключения к {connectionString}\nНастроить подключение?",
                    "Ошибка подключения",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Error
                );

                Settingscs error = new Settingscs();

                if (res == DialogResult.Yes)
                    error.ShowDialog();
                else
                    Application.Exit();
            }
        }

        // Чекбокс "Показать пароль"
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.UseSystemPasswordChar = !checkBox1.Checked;
        }

        // Ограничение ввода в поле логина — только английские символы
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.English_Symbol(sender, e);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Captcha_Load();
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            InputLimit.Captcha_Symbol(sender, e, chars);
        }

        private void Login_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                try
                {
                    BackupClass.CreateBackupWithDialog(AppDomain.CurrentDomain.BaseDirectory);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка дампа");
                    MessageBox.Show(ex.Message);
                }
                Application.Exit();
            }
        }
    }
}
