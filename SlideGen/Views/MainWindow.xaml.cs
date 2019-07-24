using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using SlideGen.Services;
using Syncfusion.Presentation;
using Brushes = System.Windows.Media.Brushes;
using Button = System.Windows.Controls.Button;
using Image = System.Windows.Controls.Image;

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
        private List<string> selectedImages;    // List of selected image URLs
        private string previousTitle;           // Previous title string value
        private List<string> keywords;          // List of keywords used in suggested image search
        private List<string> previousKeywords;  // Previous list of keywords

        public MainWindow()
        {
            InitializeComponent();
            imgService = new ImageService();
            images = new string[Constants.maxImages];
            selectedImages = new List<string>();
            InitTimer();
            previousTitle = "";
            keywords = new List<string>();
            previousKeywords = new List<string>();
            SlideBody.Document.Blocks.Add(new Paragraph(new Run("Slide text goes here")));  // set initial text for slide body
        }

        /// <summary>
        /// Initializes the timer for suggested image refresh.
        /// </summary>
        private void InitTimer()
        {
            imgTimer = new Timer();
            imgTimer.Tick += new EventHandler(imgTimer_Tick);
            imgTimer.Interval = 5000; // milliseconds
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
            // request new images if the title has changed
            if (SlideTitle.Text != previousTitle)
            {
                images = imgService.GetImages(SlideTitle.Text + " " + String.Join(" ", keywords.ToArray()));
                loadSuggestedImages(images);
                previousTitle = SlideTitle.Text;  // replace the previous title with the current one
                
            }
            else if(!keywords.SequenceEqual(previousKeywords))  //request new images if keywords have changed
            {
                images = imgService.GetImages(SlideTitle.Text + " " + String.Join(" ", keywords.ToArray()));
                loadSuggestedImages(images);
                previousKeywords.Clear();
                previousKeywords = keywords.ToList();  // replace the previous keywords with the current keywords
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
                    string bodyTxt = new TextRange(SlideBody.Document.ContentStart, SlideBody.Document.ContentEnd).Text;
                    body.TextBody.AddParagraph(bodyTxt);

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

            // for each suggested image, create a button with an image inside it and add it to the UI stackpanel
            foreach (string uri in imgURIs)
            {
                if (uri != null)
                {
                    Button imgButton = new Button
                    {
                        Height = 100,
                        Padding = new Thickness(10, 10, 10, 10),
                        Content = new Image
                        {
                            Source = new BitmapImage(new Uri(uri)),
                            VerticalAlignment = VerticalAlignment.Center
                        }
                    };
                    imgButton.Click += new RoutedEventHandler(imgClickHandler);

                    SuggestedImages.Children.Add(imgButton);
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
            if (button.Background != System.Windows.Media.Brushes.Cyan)
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

        /// <summary>
        /// Event handler for the bold text button. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bold_text(object sender, RoutedEventArgs e)
        {
            //Create a TextRange from the selection start and end and change its font weight to bold
            TextRange range = new TextRange(SlideBody.Selection.Start, SlideBody.Selection.End);
            range.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);

            //Add selected text to list of keywords
            keywordChange();
        }

        /// <summary>
        /// Wrapper for the keywordChange method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SlideBody_TextChanged(object sender, TextChangedEventArgs e)
        {
            keywordChange();
        }

        /// <summary>
        /// Clears old saved keywords and parses new ones from the bodytext.
        /// </summary>
        private void keywordChange()
        {
            //First remove all saved keywords by replacing them with null values
            keywords.Clear();

            //Move text pointer to the beginning of the text
            TextPointer position = SlideBody.Document.ContentStart;
            while (position.GetPointerContext(LogicalDirection.Forward) != TextPointerContext.Text)
            {
                position = position.GetNextContextPosition(LogicalDirection.Forward);
                if (position == null) return;
            }

            //Parse individual words from the textbox by moving pointer over text until whitespace is encountered
            while (position != SlideBody.Document.ContentEnd && position.GetPositionAtOffset(1) != null)
            {
                string currentString = "";
                int i = 1;
                while (currentString.Contains(" ") == false && position.GetPositionAtOffset(i) != null)
                {
                    TextRange range = new TextRange(position, position.GetPositionAtOffset(i));
                    currentString = range.Text;
                    ++i;
                }
                TextRange wordRange = new TextRange(position, position.GetPositionAtOffset(i-2));
                object oFont = wordRange.GetPropertyValue(Run.FontWeightProperty);
                if(oFont.ToString() == "Bold")
                {
                    keywords.Add(wordRange.Text);
                }

                position = position.GetPositionAtOffset(i - 1);
            }
        }
    }
}
