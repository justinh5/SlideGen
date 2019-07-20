using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Syncfusion.Presentation;

namespace SlideGen
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

        private void generate_Click(object sender, RoutedEventArgs e)
        {

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();


                if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    //Create a new PowerPoint presentation
                    IPresentation powerpointDoc = Presentation.Create();

                    //Add a blank slide to the presentation
                    ISlide slide = powerpointDoc.Slides.Add(SlideLayoutType.Blank);

                    //Add Title to the slide
                    IShape title = slide.AddTextBox(400, 80, 500, 100);
                    title.TextBody.AddParagraph(SlideTitle.Text);

                    //Add body text to the slide
                    IShape body = slide.AddTextBox(100, 150, 100, 200);
                    body.TextBody.AddParagraph(SlideBody.Text);

                    //Insert images on the right of the slide

                    //Save the PowerPoint presentation
                    powerpointDoc.Save(fbd.SelectedPath + "/ExampleSlide.pptx");

                    //Close the PowerPoint presentation
                    powerpointDoc.Close();
                }
            }

            //Completion message
            System.Windows.MessageBox.Show("Powerpoint slide saved!", "Powerpoint generator");
        }
    }
}
