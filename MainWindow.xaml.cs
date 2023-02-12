using Microsoft.Office.Interop.PowerPoint;
using Ookii.Dialogs.Wpf;
using System;
using System.IO;
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

        private void Browse_Click(object sender, RoutedEventArgs e)
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

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            string outPath = $@"{FolderNameTextBox.Text}\{OutName.Text}.pptx";
            string[] sourceFiles = Directory.GetFiles(FolderNameTextBox.Text, "*.pptx");

            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new();
            Presentation destPresentation = pptApplication.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

            destPresentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            foreach (string sourceFile in sourceFiles)
            {
                Presentation sourcePresentation = pptApplication.Presentations.Open(sourceFile, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse);

                destPresentation.PageSetup.SlideHeight = sourcePresentation.PageSetup.SlideHeight;
                destPresentation.PageSetup.SlideWidth = sourcePresentation.PageSetup.SlideWidth;

                var slides = sourcePresentation.Slides;
                foreach (Slide sourceSlide in slides)
                {
                    destPresentation.Slides.InsertFromFile(sourceFile, destPresentation.Slides.Count, sourceSlide.SlideNumber, sourceSlide.SlideNumber);

                    Slide destSlide = destPresentation.Slides[destPresentation.Slides.Count];
                    destSlide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
                    destSlide.Background.Fill.ForeColor.RGB = sourceSlide.Background.Fill.ForeColor.RGB;
                    destSlide.Background.Fill.BackColor.RGB = sourceSlide.Background.Fill.BackColor.RGB;
                    destSlide.ColorScheme = sourceSlide.ColorScheme;
                    NAR(sourceSlide);
                }

                NAR(slides);
                sourcePresentation?.Close();
                NAR(sourcePresentation);
            }

            destPresentation.SaveAs(outPath);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            destPresentation?.Close();
            NAR(destPresentation);
            pptApplication.Quit();
            NAR(pptApplication);
        }

        private static void NAR(object? o)
        {
            try
            {
                _ = System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null;
            }
        }
    }
}
