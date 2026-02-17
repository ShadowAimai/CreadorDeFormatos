using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Collections.Generic;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.ComponentModel;
using CreadorDeFormatos.ViewModel;
using OfficeOpenXml;
using Xceed.Words.NET;
using Xceed.Document.NET;
using Paragraph = Xceed.Document.NET.Paragraph;
using CreadorDeFormatos.Languages;

namespace CreadorDeFormatos
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MyViewModel viewModel { get; } = new MyViewModel();

        public MainWindow()
        {
            InitializeComponent();
            DataContext = viewModel;

        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
            };

            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                viewModel.ExcelFilePath = dialog.FileName;
            }

        }

        private void SelectFormatFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".docx",
                Filter = "Word Files (*.docx)|*.docx|All Files (*.*)|*.*"
            };

            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                viewModel.FormatFilePath = dialog.FileName;
            }
        }

        private void SelectOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            };

            viewModel.OutputFolderPath = dialog.ShowDialog() == CommonFileDialogResult.Ok ? dialog.FileName : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        }

        private async void GenerateFormats_Click(object sender, RoutedEventArgs e)
        {
            var progress = new Progress<int>(value =>
            {
                progressBar.Value = value;
            });

            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Visibility = Visibility.Visible;

            await Task.Run(() => GenerateFormats(progress));

            progressBar.Visibility = Visibility.Collapsed;

            MessageBox.Show(Languages.Resources.ArchiveCompleted);
        }

        private void GenerateFormats(IProgress<int> progress)
        {
            var fields = new List<string>();
            var people = new List<Dictionary<string, object>>();

            if (!File.Exists(viewModel.ExcelFilePath) || !File.Exists(viewModel.FormatFilePath))
            {
                MessageBox.Show(Languages.Resources.InvalidPaths);
                return;
            }

            try
            {

                using (var generator = new ExcelPackage(new FileInfo(viewModel.ExcelFilePath)))
                {
                    ExcelPackage.License.SetNonCommercialPersonal("Shadow");
                    var worksheet = generator.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int column = 1; column <= worksheet.Dimension.Columns; column++)
                    {
                        string fieldName = worksheet.Cells[2, column].Text;
                        if (!string.IsNullOrEmpty(fieldName))
                            fields.Add(fieldName);
                    }

                    for (int row = 3; row <= rowCount; row++)
                    {
                        bool isEmptyRow = true;
                        var person = new Dictionary<string, object>();

                        for (int column = 1; column <= fields.Count; column++)
                        {
                            string fieldName = fields[column - 1];
                            object value = worksheet.Cells[row, column].Value;

                            if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                                isEmptyRow = false;

                            person[fieldName] = value;
                        }

                        if (!isEmptyRow)
                            people.Add(person);

                        // Report progress based on row
                        int percent = (int)((double)row / rowCount * 100);
                        progress.Report(percent);
                    }
                }

                var masterDocument = DocX.Create(System.IO.Path.Combine(viewModel.OutputFolderPath,
                    "ReporteAnalabTiposSanguineos" + string.Format("{0:yyyy-MM-dd_HH-mm-ss-fff}", DateTime.Now) + ".docx"));

                int personCount = people.Count;
                int processed = 0;

                foreach (var person in people)
                {
                    using (var doc = DocX.Load(viewModel.FormatFilePath))
                    {
                        foreach (var field in fields)
                        {
                            string placeholder = $"{{{field}}}";
                            object rawValue = person[field];
                            string newValue;

                            if (rawValue is DateTime dt)
                            {
                                // Format as short date (e.g. 2/11/2026)
                                newValue = dt.ToShortDateString();

                                // Or use a custom format, e.g. "dd/MM/yyyy"
                                // newValue = dt.ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                newValue = rawValue?.ToString() ?? string.Empty;
                            }

                            var options = new StringReplaceTextOptions
                            {
                                SearchValue = placeholder,
                                NewValue = newValue,
                            };

                            doc.ReplaceText(options);
                        }
                        masterDocument.InsertDocument(doc);
                    }

                    processed++;
                    progress.Report((int)((double)processed / personCount * 100));
                }

                masterDocument.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                return;
            }

        }



    }
}