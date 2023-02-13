using Microsoft.Office.Interop.PowerPoint;
using Ookii.Dialogs.Wpf;
using System;
using System.Diagnostics;
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

        private void Merge_PPTX(object sender, RoutedEventArgs e)
        {
            string outPath = $@"{FolderNameTextBox.Text}\{OutName.Text}.pptx";
            string[] sourceFiles = Directory.GetFiles(FolderNameTextBox.Text, "*.pptx");

            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new();
            Presentation destPresentation = pptApplication.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

            destPresentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank); //Initialize the slide show to stop index error
            bool sizeSet = false;

            foreach (string sourceFile in sourceFiles)
            {
                Presentation sourcePresentation = pptApplication.Presentations.Open(sourceFile, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse);
                
                if (!sizeSet) //Fix the size of the merged slide show to the size of the first source presentation
                {
                    sizeSet = true;
                    destPresentation.PageSetup.SlideHeight = sourcePresentation.PageSetup.SlideHeight;
                    destPresentation.PageSetup.SlideWidth = sourcePresentation.PageSetup.SlideWidth;
                }

                var slides = sourcePresentation.Slides;
                foreach (Slide sourceSlide in slides)
                {
                    destPresentation.Slides.InsertFromFile(sourceFile, destPresentation.Slides.Count, sourceSlide.SlideNumber, sourceSlide.SlideNumber);

                    Slide destSlide = destPresentation.Slides[destPresentation.Slides.Count];
                   
                    /*
                     * The following is an failed attempt to fix the background not copying correctly.
                     */
                    destSlide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
                    destSlide.Background.Fill.ForeColor.RGB = sourceSlide.Background.Fill.ForeColor.RGB;
                    destSlide.Background.Fill.BackColor.RGB = sourceSlide.Background.Fill.BackColor.RGB;
                    
                    destSlide.ColorScheme = sourceSlide.ColorScheme; //Force the color scheme to be copied
                    NAR(sourceSlide); //Release the COM for the individual slide
                }

                NAR(slides); //Release the COM for the slides object
                sourcePresentation?.Close(); //Close the source
                NAR(sourcePresentation); //Release all COM for the source
            }

            destPresentation.Slides[1].Delete(); //Remove the blank initial slide
            destPresentation.SaveAs(outPath);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            destPresentation?.Close();
            NAR(destPresentation);
            pptApplication.Quit();
            NAR(pptApplication);

            ForceKill();
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

        private static void ForceKill()
        {
            Process[] powerPointProcesses = Process.GetProcessesByName("POWERPNT");
            if (powerPointProcesses.Length > 0) 
            {
                foreach (Process process in powerPointProcesses) 
                {
                    process.Kill();
                }
            }
        }
    }
}
