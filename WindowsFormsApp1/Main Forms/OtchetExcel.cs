using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class OtchetExcel : Form
    {
        string connectionString; // Строка подключения к базе данных
        DataTable orderTable; // Таблица для хранения данных о заказах

        public OtchetExcel()
        {
            InitializeComponent();
        }

        // Загрузка формы
        private void OtchetExcel_Load(object sender, EventArgs e)
        {
            try
            {
                // Устанавливаем начальные даты фильтрации
                dateFrom.Value = DateTime.Now.AddMonths(-1);
                dateTo.Value = DateTime.Now;
                dateFrom.MaxDate = DateTime.Now;
                dateTo.MaxDate = DateTime.Now.AddDays(1);

                // Получаем строку подключения
                Connect connect = new Connect();
                connectionString = connect.ConnectDB();

                // Загружаем данные из базы
                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    con.Open();
                    orderTable = new DataTable();

                    // SQL-запрос для получения информации о заказах
                    string query = @"
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
        ";

                    MySqlCommand cmd = new MySqlCommand(query, con);
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    da.Fill(orderTable);

                    // Настройка DataGridView
                    dataGridView1.DataSource = orderTable;
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

                    // Настройка ограничений даты по диапазону данных
                    if (orderTable.Rows.Count > 0)
                    {
                        var dates = orderTable.AsEnumerable()
                                              .Where(r => r["Дата"] != DBNull.Value)
                                              .Select(r => Convert.ToDateTime(r["Дата"]))
                                              .ToList();

                        if (dates.Count > 0)
                        {
                            DateTime minDate = dates.Min();
                            DateTime maxDate = dates.Max();

                            dateFrom.MinDate = minDate;
                            dateFrom.MaxDate = maxDate;
                            dateFrom.Value = minDate;

                            dateTo.MinDate = minDate;
                            dateTo.MaxDate = maxDate;
                            dateTo.Value = maxDate;
                        }
                    }
                }

                // Применяем фильтр по дате
                ApplyFilterAndSort();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных:\n" + ex.Message);
            }
        }

        // Кнопка закрытия формы
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Изменение даты начала фильтрации
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (dateFrom.Value > dateTo.Value)
                {
                    dateTo.Value = dateFrom.Value;
                }
                ApplyFilterAndSort();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при изменении даты начала фильтрации:\n" + ex.Message);
            }
        }

        // Изменение даты конца фильтрации
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (dateTo.Value < dateFrom.Value)
                {
                    dateFrom.Value = dateTo.Value;
                }
                ApplyFilterAndSort();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при изменении даты конца фильтрации:\n" + ex.Message);
            }
        }

        // Применение фильтра по дате
        private void ApplyFilterAndSort()
        {
            try
            {
                if (orderTable == null || orderTable.Rows.Count == 0)
                    return;

                DateTime start = dateFrom.Value.Date;
                DateTime end = dateTo.Value.Date.AddDays(1).AddTicks(-1); // Конец дня

                // Фильтр DataView по дате
                string filter = $"Дата >= #{start:MM/dd/yyyy}# AND Дата <= #{end:MM/dd/yyyy}#";
                DataView dv = orderTable.DefaultView;
                dv.RowFilter = filter;

                dataGridView1.DataSource = dv;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при применении фильтра:\n" + ex.Message);
            }
        }

        // Кнопка экспорта в Excel
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для отчёта.");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = $"Отчет_{dateFrom.Value:dd-MM-yyyy}_{dateTo.Value:dd-MM-yyyy}.xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            string savedFile = sfd.FileName;

            Excel.Application excel = null;
            Excel.Workbook workbook = null;

            try
            {
                excel = new Excel.Application();
                excel.DisplayAlerts = false;
                excel.Visible = false;

                string templatePath = Path.Combine(Application.StartupPath, "otchettemplate.xlsx");

                workbook = excel.Workbooks.Open(templatePath);

                Excel.Worksheet dataSheet =
                    (Excel.Worksheet)workbook.Worksheets["DATA"];

                Excel.Worksheet dashboardSheet =
                    (Excel.Worksheet)workbook.Worksheets["DASHBOARD"];

                dashboardSheet.Range["B2"].Value2 =
                    $"{dateFrom.Value:dd.MM.yyyy} - {dateTo.Value:dd.MM.yyyy}";

                Excel.ListObject table =
                    dataSheet.ListObjects["TicketsTable"];

                if (table == null)
                {
                    MessageBox.Show("Таблица TicketsTable не найдена.");
                    return;
                }

                var rows = dataGridView1.Rows
                    .Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow)
                    .ToList();

                int rowCount = rows.Count;
                int colCount = dataGridView1.Columns.Count;

                object[,] data = new object[rowCount, colCount];

                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        data[r, c] = rows[r].Cells[c].Value ?? "";
                    }
                }

                Excel.Range newRange = dataSheet.Range[
                    dataSheet.Cells[1, 1],
                    dataSheet.Cells[rowCount + 1, colCount]
                ];

                table.Resize(newRange);
                Excel.Range body = table.DataBodyRange;

                if (body != null && rowCount > 0)
                {
                    body.Value2 = data;
                }
                workbook.RefreshAll();

                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    try
                    {
                        foreach (Excel.PivotTable pt in ws.PivotTables())
                            pt.RefreshTable();
                    }
                    catch { }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                }
                workbook.SaveAs(savedFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка:\n" + ex.Message);
                return;
            }
            finally
            {
                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    }
                }
                catch { }

                try
                {
                    if (excel != null)
                    {
                        excel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                    }
                }
                catch { }

                workbook = null;
                excel = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            DialogResult result = MessageBox.Show(
                "Отчёт успешно создан!\nОткрыть файл?",
                "Готово",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = savedFile,
                    UseShellExecute = true
                });
            }
        }

        // Окрашивание строк DataGridView в зависимости от статуса
        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            try
            {
                if (dataGridView1.Rows[e.RowIndex].Cells["Статус"].Value == null)
                    return;

                string status = dataGridView1.Rows[e.RowIndex].Cells["Статус"].Value.ToString().ToLower();

                if (status.Contains("заверш"))
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightGreen;
                else if (status.Contains("отмен"))
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightCoral;
                else if (status.Contains("создан"))
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
                else
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка при окрашивании строк:\n" + ex.Message);
            }
        }
    }
}
