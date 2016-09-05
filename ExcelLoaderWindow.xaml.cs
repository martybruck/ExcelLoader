using System.Windows;
using System.Windows.Controls;

namespace ExcelLoader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ExcelLoaderViewModel dataContext = new ExcelLoaderViewModel();

        public MainWindow()
        {
            InitializeComponent();
            DataContext = dataContext;
        }
    }
}
