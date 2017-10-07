using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FileStats
{
	public struct DataRow
	{
		public DateTime SelectedDate;
		public string CarManufaturer;
		public string CarModel;
		public int CarCount;
	}

	public partial class MainWindow : Window
	{
		private BackgroundWorker _backgroundWorker
			; // объект для запуска задачи в фоновом потоке, чтобы не фризить UI, если задача будет долгой по времени

		private int _filesCount; // количество файлов в директории
		private List<DataRow> _data = new List<DataRow>(); //данные для записи в файл
		private DateTime _selectedDate;
		private int _lastExcelRow;

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
		}

		private void OnInitialized(object sender, EventArgs eventArgs)
		{
			if (!string.IsNullOrEmpty(Properties.Settings.Default.LastSelectedPath))
				PathTextBox.Text = Properties.Settings.Default.LastSelectedPath;
			if (!string.IsNullOrEmpty(Properties.Settings.Default.LastSelectedExcelFilePath))
				ExcelFilePathTextBox.Text = Properties.Settings.Default.LastSelectedExcelFilePath;
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
				if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(dialog.SelectedPath))
				{
					Properties.Settings.Default["LastSelectedPath"] = dialog.SelectedPath;
					PathTextBox.Text = dialog.SelectedPath;
					Properties.Settings.Default.Save();
				}
			}
		}

		private void BrowseExcelFileButton_OnClick(object sender, RoutedEventArgs e)
		{
			using (var dialog = new OpenFileDialog { Filter = "" })
			{
				dialog.Filter = "Excel Files(*.xlsx;*.xls)|*.xlsx;*.xls;|All files (*.*)|*.*";

				dialog.FilterIndex = 1;
				dialog.Multiselect = false;
				dialog.RestoreDirectory = true;

				if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(dialog.FileName))
				{
					Properties.Settings.Default["LastSelectedExcelFilePath"] = dialog.FileName;
					ExcelFilePathTextBox.Text = dialog.FileName;
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
				_selectedDate = DatePicker.SelectedDate ?? DateTime.Now;
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
			_backgroundWorker =
				(BackgroundWorker)FindResource("BackgroundWorker"); // взятие объекта для фоновой задачи из разметки
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
		/// 2. Пробегаемся по директориям главной папки, в каждой считаем количество файлов и рекурсивно вызываем функцию поиска,
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
				Search(directory, 0);
				int progress = (int)Math.Round((float)(i + 1) / mainDirectories.Length * 100.0f);
				_backgroundWorker.ReportProgress(progress, directory + "Files : " + _filesCount);
				_filesCount = 0;
			}
		}

		/// <summary>
		/// Рекурсивная функция перебора всех файлов
		/// </summary>
		/// <param name="mainDirectory">Директория поиска</param>
		void Search(string mainDirectory, int depth)
		{
			string[] dirs = { };
			dirs = GetDirectories(mainDirectory, dirs);
			Debug.WriteLine(depth + "\t" + mainDirectory);

			foreach (string directory in dirs)
			{
				Search(directory, depth + 1);
			}
			CountFiles(mainDirectory, depth);
			CreateDataRow(mainDirectory, depth);
		}

		private void CreateDataRow(string mainDirectory, int depth)
		{
			if (depth == 1)
			{
				Debug.WriteLine("FilesCount in " + mainDirectory + "\t" + _filesCount);
				var splits = mainDirectory.Split('\\');
				_data.Add(new DataRow
				{
					SelectedDate = _selectedDate,
					CarManufaturer = splits[splits.Length - 2],
					CarModel = splits[splits.Length - 1],
					CarCount = _filesCount
				});

				_filesCount = 0;
			}
		}

		private void CountFiles(string mainDirectory, int depth)
		{
			if (depth >= 1)
			{
				string[] files = { };
				files = GetFiles(mainDirectory, files);

				if (files.Length != 0)
				{
					_filesCount += files.Length;
				}
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
			var book = ReadExcelFile();
			var xlWorkSheet = (Excel.Worksheet)book.Worksheets.Item[1];
			_lastExcelRow = xlWorkSheet.UsedRange.Rows.Count + 1;

			if (_data.Count > 0)
			{
				foreach (var row in _data)
					AddDataToExcel(xlWorkSheet,row);
			}
			CloseExcel(book);
		}

		private Excel.Workbook ReadExcelFile()
		{
			var xlApp = new Excel.Application();
			return xlApp.Workbooks.Open(ExcelFilePathTextBox.Text, System.Reflection.Missing.Value,
				System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
				System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
				System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
				System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
				System.Reflection.Missing.Value);
		}

		public void AddDataToExcel(Excel.Worksheet worksheet, DataRow dataRow)
		{
			worksheet.Cells[_lastExcelRow, "A"] = dataRow.SelectedDate;
			worksheet.Cells[_lastExcelRow, "B"] = dataRow.CarManufaturer;
			worksheet.Cells[_lastExcelRow, "C"] = dataRow.CarModel;
			worksheet.Cells[_lastExcelRow, "D"] = dataRow.CarCount;
			_lastExcelRow++;
		}

		public void CloseExcel(Excel.Workbook workbook)
		{
			
				workbook.SaveAs(ExcelFilePathTextBox.Text, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
					System.Reflection.Missing.Value,
					System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
					System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
					System.Reflection.Missing.Value, System.Reflection.Missing.Value);


				workbook.Close(true, ExcelFilePathTextBox.Text, System.Reflection.Missing.Value); 
			
		}
	}
}
