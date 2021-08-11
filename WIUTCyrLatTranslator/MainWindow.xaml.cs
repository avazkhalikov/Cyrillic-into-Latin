using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Windows.Xps.Packaging;

namespace WIUTCyrLatTranslator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        //our global dictionary array. 
        //Important: make sure each kiril array item under the same(corresponding) latin array item index.

        char[] kiril ={'ю', 'Ю', 'я', 'Я', 'ё', 'Ё', 'ш', 'Ш', 'ч', 'Ч', 'ў', 'Ў',
                      'қ', 'Қ', 'ғ', 'Ғ', 'ц', 'Ц', 'й', 'Й', 'у', 'У', 'к', 'К',
                'е', 'Е', 'н', 'Н', 'г', 'Г', 'щ', 'Щ', 'з', 'З', 'х', 'Х',
                'э', 'Э', 'ж', 'Ж', 'д', 'Д', 'л', 'Л', 'о', 'О', 'р', 'Р',
                'п', 'П', 'а', 'А', 'в', 'В', 'ф', 'Ф', 'с', 'С', 'м', 'М',
                'и', 'И', 'т', 'Т', 'б', 'Б', 'қ', 'Қ', 'ҳ', 'Ҳ', 'ғ', 'Ғ', 'ь','ы','ъ'};

        String[] lat ={"yu", "Yu", "ya", "Ya", "yo", "Yo", "sh", "Sh", "ch", "Ch", "o'", "O'",
                      "q", "Q", "g'", "G'", "ts", "Ts", "y", "Y", "u", "U", "k", "K", "e", "E",
                "n", "N", "g", "G", "sh", "Sh", "z", "Z", "x", "X", "e", "E", "j", "J",
                "d", "D", "l", "L", "o", "O", "r", "R", "p", "P", "a", "A", "v", "V", "f",
                "F", "s", "S", "m", "M", "i", "I", "t", "T", "b", "B", "q", "Q", "h", "H",
                "g", "G", "`","y","`" };


        public MainWindow()
        {
            InitializeComponent();
        }

     

        /// <summary>
        /// This method takes a Word document full path and new XPS document full path and name
        /// and returns the new XpsDocument
        /// </summary>
        /// <param name="wordDocName"></param>
        /// <param name="xpsDocName"></param>
        /// <returns></returns>
        private XpsDocument ConvertWordDocToXPSDoc(string wordDocName, string xpsDocName)
        {
            // Create a WordApplication and add Document to it
            Microsoft.Office.Interop.Word.Application
                wordApplication = new Microsoft.Office.Interop.Word.Application();
            wordApplication.Documents.Add(wordDocName);


            Document doc = wordApplication.ActiveDocument;           
            try
            {
                doc.SaveAs(xpsDocName, WdSaveFormat.wdFormatXPS);
                wordApplication.Quit();

                XpsDocument xpsDoc = new XpsDocument(xpsDocName, System.IO.FileAccess.Read);
                return xpsDoc;
            }
            catch (Exception exp)
            {
                string str = exp.Message;
            }
            return null;
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".docx";
            dlg.Filter = "Word documents (.docx)|*.docx";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                if (dlg.FileName.Length > 0)
                {
                    SelectedFileTextBox.Text = dlg.FileName;
                    string newXPSDocumentName = String.Concat(System.IO.Path.GetDirectoryName(dlg.FileName), "\\",
                                   System.IO.Path.GetFileNameWithoutExtension(dlg.FileName), ".xps");

                    //set global filename
                    CurrentFileName = dlg.FileName;
                    // Set DocumentViewer.Document to XPS document
                    documentViewer1.Document =
                        ConvertWordDocToXPSDoc(dlg.FileName, newXPSDocumentName).GetFixedDocumentSequence();
                }
            }
        }

        /// <summary>
        /// ProgressBar make it invisible.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BGThreadWorkDone(object sender, RunWorkerCompletedEventArgs e)
        {
            pbStatus.Visibility = Visibility.Collapsed;
            

        }

        /// <summary>
        /// Convert Button Click Event Handler!
        /// </summary>
        String CurrentFileName;
        string destfileName;
        private ManualResetEvent mre = new ManualResetEvent(false);
        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            //Note: this EventHandler lines for the sake of displaying progress bar!
            //could simply call Actual_Convert()...after taking away sender, e.
            //
            pbStatus.Visibility = Visibility.Visible;     
            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += new DoWorkEventHandler(Actual_Convert);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGThreadWorkDone);

            //need to pass recieved parameter to my Actual_Convert          
            // The parameters you want to pass to the do work event of the background worker.

            var tag = sender as Button;
            string value = tag.Name; 
            worker.RunWorkerAsync(argument:value);  //sending button name as argument!

            mre.WaitOne();

            LoadNextWindow();



        }

        private void LoadNextWindow()
        {
            //creating converted xps file and loading into documentviewer2.
            string newXPSDocumentName = String.Concat(System.IO.Path.GetDirectoryName(destfileName), "\\",
                               System.IO.Path.GetFileNameWithoutExtension(destfileName), ".xps");
            documentViewer2.Document =
                    ConvertWordDocToXPSDoc(destfileName, newXPSDocumentName).GetFixedDocumentSequence();
        }

      

        /// <summary>
        /// Actual Converting happens here, in this method!
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Actual_Convert(object sender, DoWorkEventArgs e)
        {
            string value = (string)e.Argument; //get the button Name.          
            
            List<string> converting = GetDataParagraphs();

            if (converting != null && converting.Any())
            {
                //Convert Chars!
                StringBuilder sb = new StringBuilder();
                List<string> pars = new List<string>();

                if (value == "Cyrillic")
                {
                    foreach (var par in converting)
                    {
                        pars.Add(ConvertFromCyrToLatString(par));
                    }
                }
                if (value == "Latin")
                {

                    foreach (var par in converting)
                    {
                        pars.Add(ConvertFromLatToCyrString(par));
                    }
                }

                //save converted text as a word file!
                string ext = CurrentFileName.Split('.')[1];
                string name = CurrentFileName.Split('.')[0];
                ext = "_converted." + ext;
                destfileName = name + ext;

                createWordWithText(pars, destfileName);
                
                mre.Set(); //signal the WaitOne to let the next command run!

                //pop up dialog box with full link!               
                MessageBox.Show("Successfully Converted & Saved File here: " + destfileName);
            }
        }

        /// <summary>
        /// read paragraphs from word file
        /// </summary>
        /// <returns></returns>
        private List<string> GetDataParagraphs()
        {
            //read file
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Document doc = new Document();

            object fileName = CurrentFileName;
            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;
            doc = word.Documents.Open(ref fileName,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            //read parahgraph by!
            String read = string.Empty;
            List<string> data = new List<string>();
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                string temp = doc.Paragraphs[i + 1].Range.Text.Trim();
                if (temp != string.Empty)
                    data.Add(temp);
            }
                    ((_Document)doc).Close();
            ((_Application)word).Quit();
            return data;
        }

        /// <summary>
        /// save given paragraphs to given file!
        /// </summary>
        /// <param name="pars"></param>
        /// <param name="destfilename"></param>
        private void createWordWithText(List<string> pars, string destfilename)
        {
            //read file
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();            
            Document doc = new Document();          

            //adding text to document  
             doc.Content.SetRange(0, 0);
            //  doc.Content.Text = done;

            //Create a missing variable for missing value  
            foreach (var par in pars)
            {
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Paragraph para1 = doc.Content.Paragraphs.Add(ref missing);
                object styleHeading1 = "Heading 1";
                para1.Range.set_Style(ref styleHeading1);
                para1.Range.Text = par;
                para1.Range.InsertParagraphAfter();
            }

            doc.SaveAs(destfilename);

            ((_Document)doc).Close();
            ((_Application)word).Quit();


        }


      
        /// <summary>
        /// this method converts each given cyr words into latin.
        /// </summary>
        /// <param name="cyr"></param>
        /// <returns></returns>
        private string ConvertFromCyrToLatString(string cyr)
        {
            StringBuilder sb = new StringBuilder();
            char[] myChars = cyr.ToCharArray();

            foreach (var oneChar in myChars)
            {
                int foundIndex=Array.IndexOf(kiril, oneChar);
               
                if (foundIndex > -1) //exists
                {
                    //get corresponding charcter from lat.
                    string found = lat[foundIndex];
                    sb.Append(found);
                }
                else {
                    sb.Append(oneChar); //otherwise we append existing char! //could be some symbol or space.
                }

            }

            return sb.ToString(); 
        }


        /// <summary>
        /// this method converts each given Latin words into Cyrillic.
        /// </summary>
        /// <param name="cyr"></param>
        /// <returns></returns>
        private string ConvertFromLatToCyrString(string cyr)
        {
            StringBuilder sb = new StringBuilder();
            char[] myChars = cyr.ToCharArray();

            for(int i=0; i<myChars.Length; i++)
            {
                // int foundIndex = Array.IndexOf(lat, oneChar);

                var foundIndex = Array.IndexOf(lat,myChars[i].ToString()); //my second array is string, so i convert!

                //have to add logic for double letters! sh, yo, yu, ya, ts, ch, 
                if (myChars[i] == 's' && i<myChars.Length-1) 
                {
                    //check if next one is h. or split by 2 chars!
                    if (myChars[i + 1] == 'h')
                    {
                        
                        foundIndex = Array.IndexOf(lat, "sh"); //my second array is string, so i convert!
                        i++; //this forces loop not to evaluate the next one!
                    }
                }
                if (myChars[i] == 'y' && i < myChars.Length - 1)
                {
                    //check if next one is h. or split by 2 chars!
                    if (myChars[i + 1] == 'u')
                    {

                        foundIndex = Array.IndexOf(lat, "yu"); //my second array is string, so i convert!
                        i++; //this forces loop not to evaluate the next one!
                    }
                }
                if (myChars[i] == 'y' && i < myChars.Length - 1)
                {
                    //check if next one is h. or split by 2 chars!
                    if (myChars[i + 1] == 'a')
                    {

                        foundIndex = Array.IndexOf(lat, "ya"); //my second array is string, so i convert!
                        i++; //this forces loop not to evaluate the next one!
                    }
                }
                if (myChars[i] == 't' && i < myChars.Length - 1)
                {
                    //check if next one is h. or split by 2 chars!
                    if (myChars[i + 1] == 's')
                    {

                        foundIndex = Array.IndexOf(lat, "ts"); //my second array is string, so i convert!
                        i++; //this forces loop not to evaluate the next one!
                    }
                }

                if (myChars[i] == 'c' && i < myChars.Length - 1)
                {
                    //check if next one is h. or split by 2 chars!
                    if (myChars[i + 1] == 'h')
                    {

                        foundIndex = Array.IndexOf(lat, "ch"); //my second array is string, so i convert!
                        i++; //this forces loop not to evaluate the next one!
                    }
                }


                if (foundIndex > -1) //exists
                {
                    //get corresponding charcter from lat.
                    char found = kiril[foundIndex];
                    sb.Append(found);
                }
                else
                {
                    sb.Append(myChars[i]); //otherwise we append existing char! //could be some symbol or space.
                }               

            }

            return sb.ToString();
        }



        private void Save_Click_1(object sender, RoutedEventArgs e)
        {

        }

      
    }
}
