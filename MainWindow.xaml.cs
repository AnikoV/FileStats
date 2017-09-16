using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Forms;

namespace FileStats
{
    public partial class MainWindow : Window
    {
        private BackgroundWorker _backgroundWorker; // объект для запуска задачи в фоновом потоке, чтобы не фризить UI, если задача будет долгой по времени
        private int _filesCount; // количество файлов в директории
        private Dictionary<string, int> _dataDictionary; //данные для записи в файл

        /*----------Обработка жизненного цикла окна------------*/
        
         /// <summary>
         /// Конструктор окна
         /// </summary>
        public MainWindow()
        {
            InitializeComponent(); // инициализация окна
            Initialized += OnInitialized; // подписка на событие по окончании инициализации окна
            Closed += OnClosed; // подписка на событие закрытия окна
            InitBackgroundWorker();
            _dataDictionary = new Dictionary<string, int>();
        }

        private void OnInitialized(object sender, EventArgs eventArgs)
        {
            if (!string.IsNullOrEmpty(Properties.Settings.Default.LastSelectedPath))
            {
                PathTextBox.Text = Properties.Settings.Default.LastSelectedPath;
            }
        }

        private void OnClosed(object sender, EventArgs eventArgs)
        {
            Closed -= OnClosed;
            _backgroundWorker.ProgressChanged -= BW_ProgressChanged;
            _backgroundWorker.DoWork -= BW_DoWork;
            _backgroundWorker.RunWorkerCompleted -= BW_RunWorkerCompleted;
            _backgroundWorker.Dispose();
        }
        
        /*------------------------------------------------------------------------*/


        /*----------Обработка событий взаимодействия пользователя с UI------------*/

        private void Browse_click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    Properties.Settings.Default["LastSelectedPath"] = dialog.SelectedPath;
                    PathTextBox.Text = dialog.SelectedPath;
                    Properties.Settings.Default.Save();
                }
            }
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_backgroundWorker.WorkerSupportsCancellation)
                _backgroundWorker.CancelAsync();
        }

        private void DoActionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(PathTextBox.Text))
            {
                _backgroundWorker.RunWorkerAsync(PathTextBox.Text);
            }
        }

        /*------------------------------------------------------------------------*/


        /*----------Обработка жихненного цикла фоновой задачи------------*/

        /// <summary>
        /// Инициализация фоновой задачи
        /// </summary>
        private void InitBackgroundWorker()
        {
            _backgroundWorker = (BackgroundWorker)FindResource("BackgroundWorker"); // взятие объекта для фоновой задачи из разметки
            _backgroundWorker.WorkerReportsProgress = true; // объект поддерживает прогресс
            _backgroundWorker.WorkerSupportsCancellation = true; // объект поддерживает отмену
            _backgroundWorker.ProgressChanged += BW_ProgressChanged; // подписка на событие изменения прогресса
            _backgroundWorker.DoWork += BW_DoWork; // подписка на событие выполнения задачи
            _backgroundWorker.RunWorkerCompleted += BW_RunWorkerCompleted; // подписка на событие окончания задачи
        }

        /// <summary>
        /// Функция обработчик выполнения фоновой задачи
        /// </summary>
        /// <param name="sender">объект-отправитель события</param>
        /// <param name="doWorkEventArgs">данные для выполнения задачи</param>
        private void BW_DoWork(object sender, DoWorkEventArgs doWorkEventArgs)
        {
            CountFilesInDirectory((string)doWorkEventArgs.Argument);
        }

        /// <summary>
        /// Обработчик события изменения прогресса фоновой задачи
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="progressChangedEventArgs"></param>
        private void BW_ProgressChanged(object sender, ProgressChangedEventArgs progressChangedEventArgs)
        {
            ProgressBar.Value = progressChangedEventArgs.ProgressPercentage;
            if (progressChangedEventArgs.UserState != null)
                TextBox.Text += "\n" + progressChangedEventArgs.UserState;
        }

        /// <summary>
        /// Обработчик события окончания задачи
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="runWorkerCompletedEventArgs"></param>
        private void BW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs runWorkerCompletedEventArgs)
        {
            WriteResultInCsvFile();
        }

        /*------------------------------------------------------------------------*/

        /*----------Работа с файлами и папками------------*/

        /// <summary>
        /// Функция пересчета файлов и директорий в указанной (главной) директории
        /// 1. Считываем директории и кол-во файлов в главной папке
        /// 2. Пробегаемся по директориям главной папки, в каждой считаем колличество файлов и рекурсивно вызываем функцию поиска,
        ///  чтобы пройтись в глубину директории, пока она не закончится
        /// </summary>
        /// <param name="rootDirectory">Путь к папке</param>
        void CountFilesInDirectory(string rootDirectory)
        {
            //1
            var mainDirectories = Directory.GetDirectories(rootDirectory);
            var filesInRootDirectory = Directory.GetFiles(rootDirectory).Length;
            _backgroundWorker.ReportProgress(0, rootDirectory + " Files : " + filesInRootDirectory);

            //2
            for (var i = 0; i < mainDirectories.Length; i++)
            {
                string directory = mainDirectories[i];
                string[] files = { };

                files = GetFiles(directory, files);
                if (files.Length != 0)
                {
                    _filesCount += files.Length;
                }
                Search(directory);
                int progress = (int)Math.Round((float)(i + 1) / mainDirectories.Length * 100.0f);
                _dataDictionary[directory] = _filesCount;
                _backgroundWorker.ReportProgress(progress, directory + "Files : " + _filesCount);
                _filesCount = 0;
            }
        }

        /// <summary>
        /// Рекурсивная функция перебора всех файлов
        /// </summary>
        /// <param name="mainDirectory">Директория поиска</param>
        void Search(string mainDirectory)
        {
            string[] dirs = { };
            dirs = GetDirectories(mainDirectory, dirs);

            foreach (string directory in dirs)
            {
                string[] files = { };
                files = GetFiles(directory, files);

                if (files.Length != 0)
                {
                    _filesCount += files.Length;
                }
                Search(directory);
            }
        }

        /// <summary>
        /// Взятие всех директорий внутри указанной с обработкой исключений 
        /// (например, если нет доступа, чтобы программа не упала с ошибкой, 
        /// а например вывела все папки к которым не получила доступ)
        /// </summary>
        /// <param name="mainDirectory">Дректоррия поиска</param>
        /// <param name="dirs">массив для записи результата</param>
        /// <returns></returns>
        private string[] GetDirectories(string mainDirectory, string[] dirs)
        {
            try
            {
                dirs = Directory.GetDirectories(mainDirectory);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return dirs;
        }

        /// <summary>
        /// Взятие всех файлов внутри указанной директории с обработкой исключений 
        /// (например, если нет доступа к файлу, чтобы программа не упала с ошибкой, 
        /// а например вывела все файлы к которым не получила доступ)
        /// </summary>
        /// <param name="directory">Дректоррия поиска</param>
        /// <param name="files">массив для записи результата</param>
        /// <returns></returns>
        private string[] GetFiles(string directory, string[] files)
        {
            try
            {
                files = Directory.GetFiles(directory);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }

            return files;
        }

        /*------------------------------------------------------------------------*/
        
        /// <summary>
        /// Запись в файл в формате CSV
        /// </summary>
        private void WriteResultInCsvFile()
        {
            if (_dataDictionary.Count > 0)
            {
                using (var file = new StreamWriter("output.csv"))
                {
                    foreach (var kv in _dataDictionary)
                    {
                        file.WriteLine(string.Format("{0};{1}", kv.Key, kv.Value));
                    }
                }
            }
        }
    }
}
