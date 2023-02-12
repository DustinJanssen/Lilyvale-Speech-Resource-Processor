using Ookii.Dialogs.Wpf;
using System.Windows;

namespace Lilyvale_Speech_Resource_Processor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog()
            {
                Description = "Select Folder",
                RootFolder = System.Environment.SpecialFolder.Desktop,
                ShowNewFolderButton= true,
                UseDescriptionForTitle= true
            };

            var result = dialog.ShowDialog();
            if (result.HasValue && result.Value)
            {
                string folderPath = dialog.SelectedPath;
                FolderNameTextBox.Text = folderPath;

            }
        }
    }
}
