using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    static class BackupClass
    {
        private static void CreateDatabaseDump(string fullPath)
        {
            Connect connect = new Connect();
            using (MySqlConnection conn = new MySqlConnection(connect.ConnectDB()))
            {
                using (MySqlCommand cmd = new MySqlCommand())
                {
                    using (MySqlBackup mb = new MySqlBackup(cmd))
                    {
                        cmd.Connection = conn;
                        conn.Open();
                        mb.ExportToFile(fullPath);
                        conn.Close();
                    }
                }
            }
        }

        public static string CreateBackupWithDialog(string selectedFolder)
        {
            if (!System.IO.Directory.Exists(selectedFolder))
            {
                throw new DirectoryNotFoundException($"Папка не найдена: {selectedFolder}");
            }

            string folderName = $"dump_{DateTime.Now:yyyy_MM_dd_HH_mm_ss}";
            string backupDirectory = Path.Combine(selectedFolder, folderName);

            System.IO.Directory.CreateDirectory(backupDirectory);

            string fileName = $"dump_{DateTime.Now:yyyy_MM_dd_HH_mm_ss}.sql";
            string fullPath = Path.Combine(backupDirectory, fileName);

            CreateDatabaseDump(fullPath);

            // Добавляем CREATE DATABASE и USE в начало дампа
            string dumpContent = File.ReadAllText(fullPath, Encoding.UTF8);

            string header =
                                @"CREATE DATABASE IF NOT EXISTS db71
                        CHARACTER SET utf8mb4
                        COLLATE utf8mb4_unicode_ci;

                        USE db71;

                        ";

            File.WriteAllText(fullPath, header + dumpContent, Encoding.UTF8);

            return fullPath;
        }

        public static void RestoreBackup(string sqlFilePath)
        {
            if (!File.Exists(sqlFilePath))
            {
                throw new FileNotFoundException("Файл резервной копии не найден.");
            }

            if (new FileInfo(sqlFilePath).Length == 0)
                throw new Exception("Файл резервной копии пустой.");

            Connect connect = new Connect();

            using (MySqlConnection conn = new MySqlConnection(connect.ConnectNoDB()))
            {
                using (MySqlCommand cmd = new MySqlCommand())
                {
                    using (MySqlBackup mb = new MySqlBackup(cmd))
                    {
                        cmd.Connection = conn;

                        conn.Open();

                        mb.ImportFromFile(sqlFilePath);

                        conn.Close();
                    }
                }
            }
        }
    }
}
