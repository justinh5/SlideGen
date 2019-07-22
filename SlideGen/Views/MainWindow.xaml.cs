using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using SlideGen.Services;
using Syncfusion.Presentation;
using Button = System.Windows.Controls.Button;

namespace SlideGen
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Timer imgTimer;                 // Timer for suggested image refresh
        private ImageService imgService;        // Image retrieval service
        private string[] images;                // Array of retrieved image URLs (small size)
        private List<string> selectedImages;    // Array of selected image URLs
        private string previousTitle;           // Previous title string value

        public MainWindow()
        {
            InitializeComponent();
            imgService = new ImageService();
            images = new string[Constants.maxImages];
            selectedImages = new List<string>();
            InitTimer();
            previousTitle = "";
        }

        /// <summary>
        /// Initializes the timer for suggested image refresh.
        /// </summary>
        public void InitTimer()
        {
            imgTimer = new Timer();
            imgTimer.Tick += new EventHandler(imgTimer_Tick);
            imgTimer.Interval = 5000; // miliseconds
            imgTimer.Start();
        }

        /// <summary>
        /// Invoke every 5 seconds on timer tick. Retrieve new images using the image service if the title has 
        /// changed since the last tick, and udate the suggested image area.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void imgTimer_Tick(object sender, EventArgs e)
        {
            // only make image request if the title has changed
            if (SlideTitle.Text != previousTitle)
            {
                images = imgService.GetImages(SlideTitle.Text);
                loadSuggestedImages(images);
                previousTitle = SlideTitle.Text;  // replace the previous title with the current one
            }
        }

        /// <summary>
        /// Event handler to generate a single powerpoint slide. The Title, body text, and 
        /// all selected images are included on the slide. Images must be downloaded from the source
        /// to the selected folder.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void generate_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();  //User selects folder location to save powerpoint

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
                    IShape body = slide.AddTextBox(100, 150, 500, 200);
                    body.TextBody.AddParagraph(SlideBody.Text);

                    //Insert images in the slide on the right
                    int i = 0, top = 100;
                    foreach (string uri in selectedImages)
                    {
                        string imgPath = "\\image" + i.ToString() + ".jpg";

                        //Download images before inserting in slide
                        using (WebClient client = new WebClient())
                        {
                            client.DownloadFile(new Uri(uri), fbd.SelectedPath + imgPath);
                        }

                        //Insert images on the right of the slide
                        Stream picStream = File.Open(fbd.SelectedPath + imgPath, FileMode.Open);

                        //Adds the picture to a slide by specifying its size and position.
                        IPicture picture = slide.Pictures.AddPicture(picStream, 600, top, 100, 100);

                        //Dispose the image stream
                        picStream.Dispose();

                        ++i; top += 100;
                    }

                    //Save the PowerPoint presentation
                    powerpointDoc.Save(fbd.SelectedPath + "/ExampleSlide.pptx");

                    //Close the PowerPoint presentation
                    powerpointDoc.Close();
                }
            }

            //Completion message
            System.Windows.MessageBox.Show("Powerpoint slide saved!", "Powerpoint generator");
        }

        /// <summary>
        /// Loads the images retrieved from the image service to the UI. The images are the contents of 
        /// buttons, which are clickable to toggle a seletion on or off.
        /// </summary>
        /// <param name="imgURIs">Array of image URLs</param>
        private void loadSuggestedImages(string[] imgURIs)
        {
            //First clear all saved selected images and displayed images
            selectedImages.Clear();   
            SuggestedImages.Children.Clear();

            foreach (string uri in imgURIs)
            {
                if (uri != null)
                {
                    StackPanel sp = new StackPanel();
                    Button imgButton = new Button
                    {
                        Height = 100,
                        Content = new Image
                        {
                            Source = new BitmapImage(new Uri(uri)),
                            VerticalAlignment = VerticalAlignment.Center
                        }
                    };
                    imgButton.Click += new RoutedEventHandler(imgClickHandler);

                    sp.Children.Add(imgButton);
                    SuggestedImages.Children.Add(sp);
                }
            }
        }

        /// <summary>
        /// Handler for an image's click event. The background of the image is toggled Cyan if selected
        /// or default grey when unselected.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void imgClickHandler(object sender, RoutedEventArgs e)
        {
            // Selected button and image
            Button button = (Button)sender;
            Image image = (Image)((Button)sender).Content;

            // Highlight selected image and add image url to selected list 
            if (button.Background != Brushes.Cyan)
            {
                button.Background = Brushes.Cyan;
                selectedImages.Add(image.Source.ToString());

            }
            else    // Remove highlight from selected image and remove image url from selected list
            {
                button.Background = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FFDDDDDD"));
                selectedImages.Remove(image.Source.ToString());
            }
        }
    }
}
