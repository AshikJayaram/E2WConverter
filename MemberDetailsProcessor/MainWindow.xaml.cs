using DocumentProcessor.ExcelProcessor;
using DocumentProcessor.WordProcessor;
using System;
using System.IO;
using System.Threading;
using System.Windows;

namespace MemberDetailsProcessor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            convert.IsEnabled = false;
            textBlock.Text = "Upload an excel file";
            //this.Background = new ImageBrush(new BitmapImage(new Uri(@"D:\\NodeProjects\\WebApp\\MemberDetailsProcessor\\wp.jpg")));
        }

        private const string ExcelContentType_2007 = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        private const string ExcelContentType_2003 = "application/vnd.ms-excel";

        private string UploadPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ("Uploads"));

        private Microsoft.Win32.OpenFileDialog openFileDialog = null;

        private void upload_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog = new Microsoft.Win32.OpenFileDialog();

            if (openFileDialog.ShowDialog() == true)
            {
                if (openFileDialog.CheckFileExists &&
                    (this.GetMimeType(openFileDialog.SafeFileName) == ExcelContentType_2007 
                        || this.GetMimeType(openFileDialog.SafeFileName) == ExcelContentType_2003))
                {
                    textBox.Text = openFileDialog.SafeFileName;

                    bool fileCopyCheck = this.UploadFileToFolder(openFileDialog);
                    if (fileCopyCheck)
                    {
                        //MessageBox.Show("File copied! Proceed to convert");
                        textBlock.Text = "File uploaded! Proceed to convert";
                        convert.IsEnabled = true;
                    }
                    else
                    {
                        //MessageBox.Show("Coudn't copy the file. Please try again");
                        textBlock.Text = "Coudn't upload the file. Please try again";
                    }
                }
                else
                {
                    MessageBox.Show("Please upload Excel file only");
                }
            }
        }

        private void convert_Click(object sender, RoutedEventArgs e)
        {
            IExcelReader excelReader = new ExcelReader();

            var memberList = excelReader.GetMembersList(openFileDialog.SafeFileName);

            Thread.Sleep(100);

            IDocumentWriter docWriter = new DocumentWriter();

            docWriter.WriteToDocument(memberList);

            textBlock.Text = "Document write completed";
        }

        private bool UploadFileToFolder(Microsoft.Win32.OpenFileDialog openFileDialog)
        {
            if (!Directory.Exists("Uploads"))
                Directory.CreateDirectory("Uploads");

            if (File.Exists(UploadPath + "\\" + openFileDialog.SafeFileName))
                File.Delete(UploadPath + "\\" + openFileDialog.SafeFileName);

            File.Copy(openFileDialog.FileName, UploadPath + "\\" + openFileDialog.SafeFileName);

            if (File.Exists(UploadPath + "\\" + openFileDialog.SafeFileName))
                return true;
            else
                return false;
        }

        private string GetMimeType(string fileName)
        {
            string mimeType = "application/unknown";
            string ext = System.IO.Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
            if (regKey != null && regKey.GetValue("Content Type") != null)
                mimeType = regKey.GetValue("Content Type").ToString();
            return mimeType;
        }
    }
}