using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WPFtoPTTLibrary;
using Powerpoint = Microsoft.Office.Interop.PowerPoint;

namespace PPT_Creator
{
    public partial class MainWindow : Window
    {
        private List<SlideModel> slideComponents = new List<SlideModel>();
        private static int slideCount = 0;


        public MainWindow()
        {
            InitializeComponent();
            //used to start HTTPClient and make sure there is only 1 open per application
            ApiHelper.InitilizeClient();
        }

        private void BuildPTT()
        {
            Powerpoint.Application pptApplication = new Powerpoint.Application();
            Presentation pptPresentation = pptApplication.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
            Slide slide;
            Slides slides;
            Powerpoint.TextRange title;
            Powerpoint.TextRange content;
            Powerpoint.Shape contentBox;
            CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[2];
            for (int i = 0; i < slideComponents.Count; i++)
            {
                slides = pptPresentation.Slides;
                slide = slides.AddSlide(i + 1, customLayout);
                title = slide.Shapes[1].TextFrame.TextRange;
                title.Text = $"{slideComponents[i].Title}";
                contentBox = slide.Shapes[2];
                contentBox.Width = contentBox.Width / 2;
                content = slide.Shapes[2].TextFrame.TextRange;
                content.Text = slideComponents[i].Content;
                for (int j = 0; j < slideComponents[i].Images.Count(); j++)
                {
                    slide.Shapes.AddPicture(slideComponents[i].Images[j], Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 500, 25 + (j * 100), 160, 90);
                }
            }
            slideComponents.Clear();
            slideCount = 0;
            slideLbl.Content = slideCount.ToString();
            BuildPPTButton.IsEnabled = false;
            MessageBox.Show("Your PowerPoint has been created");
        }

        string StringFromRichTextBox(RichTextBox rtb)
        {
            System.Windows.Documents.TextRange textRange = new System.Windows.Documents.TextRange(
                rtb.Document.ContentStart,
                rtb.Document.ContentEnd
            );
            return textRange.Text;
        }

        private async Task LoadImages(string searchTerms)
        {
            List<Image> ImageAreas = new List<Image>() { Image1, Image2, Image3, Image4, Image5, Image6, Image7, Image8 };
            var images = await ImageProcessor.LoadImages(searchTerms);
            for (int i = 0; i < images.Items.Count(); i++)
            {
                string imageLink = images.Items[i].Link;
                var uriSource = new Uri(imageLink, UriKind.Absolute);
                ImageAreas[i].Source = new BitmapImage(uriSource);
            }
        }

        private async void GetSearchTerms()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(TitleTextBox.Text);
            var boldItems = ContentTextBox.Document.Blocks
                .OfType<Paragraph>()
                .SelectMany(x => x.Inlines.OfType<Run>().Where(y => y.FontWeight.ToOpenTypeWeight() > 400))
                .Select(x => x.Text);
            foreach (var item in boldItems)
            {
                sb.Append(" " + item);
            }
            await LoadImages(sb.ToString());
            
        }

        private void GetImagesButton_Click(object sender, RoutedEventArgs e)
        {
             GetSearchTerms();
        }

        //Adds slide to slide components and clears and updates GIU to let user know the slide was added
        private void AddSlideButton_Click(object sender, RoutedEventArgs e)
        {
            List<Image> ImageAreas = new List<Image>() { Image1, Image2, Image3, Image4, Image5, Image6, Image7, Image8 };
            List<string> selectedImages = new List<string>();
            foreach (var image in ImageListBox.SelectedItems)
            {
                for (int i = 0; i < ImageAreas.Count; i++)
                {
                    if (image == ImageAreas[i])
                    {
                        selectedImages.Add(ImageAreas[i].Source.ToString());
                    }
                }
            }

            SlideModel slide = new SlideModel();
            foreach (string imgPath in selectedImages)
            {
                slide.Images.Add(imgPath);
            }
            slide.Title = TitleTextBox.Text;
            slide.Content = StringFromRichTextBox(ContentTextBox);
            slideComponents.Add(slide);
            selectedImages.Clear();
            ImageListBox.UnselectAll();
            BuildPPTButton.IsEnabled = true;
            TitleTextBox.Clear();
            ContentTextBox.Document.Blocks.Clear();
            slideCount++;
            slideLbl.Content = slideCount.ToString();
            
        }

        //Prevents user from building a PowerPoint with no slides by disabling button before user adds at least 1 slide
        private void BuildPPTButton_Click(object sender, RoutedEventArgs e)
        {
            BuildPTT();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            BuildPPTButton.IsEnabled = false;
            slideLbl.Content = slideCount.ToString();
        }

    }
}
