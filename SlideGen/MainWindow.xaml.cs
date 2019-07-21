using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using SlideGen.Services;
using Syncfusion.Presentation;

namespace SlideGen
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Timer imgTimer;
        private ImageService imgService;
        private string previousTitle;

        public MainWindow()
        {
            InitializeComponent();
            imgService = new ImageService();
            InitTimer();
            previousTitle = "";
        }
        public void InitTimer()
        {
            imgTimer = new Timer();
            imgTimer.Tick += new EventHandler(imgTimer_Tick);
            imgTimer.Interval = 5000; // miliseconds
            imgTimer.Start();
        }

        private void imgTimer_Tick(object sender, EventArgs e)
        {
            // only make image request if the title has changed
            if (SlideTitle.Text != previousTitle)
            {
                string[] results = imgService.GetImages(SlideTitle.Text);
                loadSuggestedImages(results);
                previousTitle = SlideTitle.Text;  // replace the previous title with the current one
            }
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
                    TextRange storedTextContent = new TextRange(SlideBody.Document.ContentStart, SlideBody.Document.ContentEnd);
                    string bodyTxt = storedTextContent.Text;
                    body.TextBody.AddParagraph(bodyTxt);

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


        private void loadSuggestedImages(string[] imgURIs)
        {
            SuggestedImages.Children.Clear();
            foreach (string uri in imgURIs)
            {
                StackPanel sp = new StackPanel();
                Image thumb = new Image();
                thumb.Height = 100;
                BitmapImage BitImg = new BitmapImage(new Uri(uri));
                thumb.Source = BitImg;
                System.Windows.Controls.CheckBox cb = new System.Windows.Controls.CheckBox();
                cb.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;

                sp.Children.Add(thumb);
                sp.Children.Add(cb);

                SuggestedImages.Children.Add(sp);
            }
        }
    }
}
