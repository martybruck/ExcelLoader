using System.Collections.Generic;
using System.ComponentModel;
using ExcelLoader.Helpers;
using System.Windows;
using System.Windows.Input;
using System;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using Microsoft.Win32;
using System.IO;

namespace ExcelLoader
{
    public class ExcelLoaderViewModel : INotifyPropertyChanged
    {

        private static ExcelLoaderViewModel _instance;

        public static ExcelLoaderViewModel GetInstance()
        {
            return _instance;
        }
        public ExcelLoaderViewModel()
        {
            _instance = this;
            MyExcelHelper = new ExcelHelper();
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(CloseWorkbook);
        }

        #region Properties

        public bool IsLoaded
        {
            get
            {
                return (ExcelFileName != null);
            }
        }

        string excelFileName = "";
        public string ExcelFileName
        {
            get
            {
                return excelFileName;
            }
            set
            {
                excelFileName = value;
                NotifyPropertyChanged();
            }
        }
        private string currentStatus = "None";
        public string CurrentStatus
        {
            get
            {
                return currentStatus;
            }
            set
            {
                currentStatus = value;
                NotifyPropertyChanged();
            }
        }

        private ObservableCollection<RuleDefinition> rules ;

        public ObservableCollection<RuleDefinition> Rules
        {
            get
            {
                return rules;
            }
            set
            {
                rules = value;
                NotifyPropertyChanged();
            }
        }
        private ExcelHelper MyExcelHelper { get; set; }

        #endregion Properties


        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property.
        // The CallerMemberName attribute that is applied to the optional propertyName
        // parameter causes the property name of the caller to be substituted as an argument.
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private RelayCommand loadExcelFileCommand;
        public ICommand LoadExcelFileCommand
        {
            get
            {
                return loadExcelFileCommand ?? (loadExcelFileCommand = new RelayCommand(p => LoadExcelFile()));
            }
        }


        public void LoadExcelFile()
        {
            var ofd = new OpenFileDialog();
            ofd.Filter = "Excel files (*.xlsx) | *.xlsx;";
            ofd.Multiselect = false;
            var result = ofd.ShowDialog();
            if (result == true)
            {
                ExcelFileName = ofd.FileName;
                CurrentStatus = "Loading file: " + ExcelFileName;
                MyExcelHelper.Load(ExcelFileName);
                Rules = MyExcelHelper.Rules;
                CurrentStatus = "File Loaded";
            }                        
        }

        private RelayCommand generateExcelResultsCommand;
        public ICommand GenerateExcelResultsCommand
        {
            get
            {
                return generateExcelResultsCommand ?? (generateExcelResultsCommand = new RelayCommand(p => GenerateExcelResults()));
            }
        }

        public void GenerateExcelResults()
        {
            if (ExcelFileName == null)
            {
                MessageBox.Show("Please Load a file before attempting to generate the results", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                CurrentStatus = "Generating Results";
                MyExcelHelper.GenerateOutputSheets(ExcelFileName);
                CurrentStatus = "Result Generation Completed - File Cleared";
                ExcelFileName = null;
                Rules = null;
                MyExcelHelper = new ExcelHelper();
            }
        }
        private void CloseWorkbook(object sender, EventArgs e)
        {
            MyExcelHelper.CloseWorkbook();
        }

    }
}
