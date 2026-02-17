using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreadorDeFormatos.ViewModel
{
    public class MyViewModel : INotifyPropertyChanged
    {

        private string _excelFilePath;
        public string ExcelFilePath
        {
            get => _excelFilePath;
            set
            {
                if (_excelFilePath != value)
                {
                    _excelFilePath = value;
                    OnPropertyChanged(nameof(ExcelFilePath));
                }

            }
        }

        private string _formatFilePath;
        public string FormatFilePath
        {
            get => _formatFilePath;
            set
            {
                if (_formatFilePath != value)
                {
                    _formatFilePath = value;
                    OnPropertyChanged(nameof(FormatFilePath));
                }

            }
        }

        private string _outputFolderPath;
        public string OutputFolderPath
        {
            get => _outputFolderPath;
            set
            {
                if (_outputFolderPath != value)
                {
                    _outputFolderPath = value;
                    OnPropertyChanged(nameof(OutputFolderPath));
                }

            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName) { 
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName)); 
        }
    }
}
