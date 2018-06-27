using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using VaderSharp;


namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {


        //initialize the space for our dictionary data
        DictionaryData DictData = new DictionaryData();



        //this is what runs at initialization
        public Form1()
        {

            InitializeComponent();

            foreach(var encoding in Encoding.GetEncodings())
            {
                EncodingDropdown.Items.Add(encoding.Name);
            }

            try
            {
                EncodingDropdown.SelectedIndex = EncodingDropdown.FindStringExact("utf-8");
            }
            catch
            {
                EncodingDropdown.SelectedIndex = EncodingDropdown.FindStringExact(Encoding.Default.BodyName);
            }
            


        }







        private void StartButton_Click(object sender, EventArgs e)
        {


            

                    FolderBrowser.Description = "Please choose the location of your .txt files to analyze";
                    if (FolderBrowser.ShowDialog() != DialogResult.Cancel) {

                        DictData.TextFileFolder = FolderBrowser.SelectedPath.ToString();
                
                        if (DictData.TextFileFolder != "")
                        {

                            saveFileDialog.FileName = "VADER-Tots.csv";

                            saveFileDialog.InitialDirectory = DictData.TextFileFolder;
                            if (saveFileDialog.ShowDialog() != DialogResult.Cancel) {


                                DictData.OutputFileLocation = saveFileDialog.FileName;

                                if (DictData.OutputFileLocation != "") {


                                    StartButton.Enabled = false;
                                    ScanSubfolderCheckbox.Enabled = false;
                                    EncodingDropdown.Enabled = false;
                            
                                    BgWorker.RunWorkerAsync(DictData);
                                }
                            }
                        }

                    }

                

        }






        private void BgWorkerClean_DoWork(object sender, DoWorkEventArgs e)
        {


            DictionaryData DictData = (DictionaryData)e.Argument;

            SentimentIntensityAnalyzer VADER = new SentimentIntensityAnalyzer();

            //set up our sentence boundary detection
            Regex SentenceSplitter = new Regex(@"(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|\!)\s", RegexOptions.Compiled);

            //selects the text encoding based on user selection
            Encoding SelectedEncoding = null;
            this.Invoke((MethodInvoker)delegate ()
            {
                SelectedEncoding = Encoding.GetEncoding(EncodingDropdown.SelectedItem.ToString());
            });



            //get the list of files
            var SearchDepth = SearchOption.TopDirectoryOnly;
            if (ScanSubfolderCheckbox.Checked)
            {
                SearchDepth = SearchOption.AllDirectories;
            }
            var files = Directory.EnumerateFiles(DictData.TextFileFolder, "*.txt", SearchDepth);



            try {

            //open up the output file
            using (StreamWriter outputFile = new StreamWriter(new FileStream(DictData.OutputFileLocation, FileMode.Create), SelectedEncoding))
            {

                using (StreamWriter outputFileSentences = new StreamWriter(new FileStream(AddSuffix(DictData.OutputFileLocation, "_Sentences"), FileMode.Create), SelectedEncoding))
                {


                    //write the header row to the output file
                    StringBuilder HeaderString = new StringBuilder();
                    HeaderString.Append("\"Filename\",\"WC\",\"Sentences\",\"Classification\",\"Compound_M\",\"Positive_M\",\"Negative_M\",\"Neutral_M\"");
                    outputFile.WriteLine(HeaderString.ToString());

                    StringBuilder HeaderStringSentence = new StringBuilder();
                    HeaderStringSentence.Append("\"Filename\",\"WC\",\"Sentence\",\"Classification\",\"Compound_M\",\"Positive_M\",\"Negative_M\",\"Neutral_M\"");
                    outputFileSentences.WriteLine(HeaderStringSentence.ToString());

                    foreach (string fileName in files)
                    {

                        //set up our variables to report
                        string Filename_Clean = Path.GetFileName(fileName);
                        Dictionary<string, int> DictionaryResults = new Dictionary<string, int>();

                        //report what we're working on
                        FilenameLabel.Invoke((MethodInvoker)delegate
                        {
                            FilenameLabel.Text = "Analyzing: " + Filename_Clean;
                        });




                        //read in the text file, convert everything to lowercase
                        string InputText = File.ReadAllText(fileName, SelectedEncoding).Trim();





                        string[] Sentences = SentenceSplitter.Split(InputText).Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();

                        int TotalStringLength = InputText.Split().Where(x => !string.IsNullOrWhiteSpace(x)).ToArray().Length;
                        int TotalSentences = Sentences.Length;



                        //     _                _                 _____         _   
                        //    / \   _ __   __ _| |_   _ _______  |_   _|____  _| |_ 
                        //   / _ \ | '_ \ / _` | | | | |_  / _ \   | |/ _ \ \/ / __|
                        //  / ___ \| | | | (_| | | |_| |/ /  __/   | |  __/>  <| |_ 
                        // /_/   \_\_| |_|\__,_|_|\__, /___\___|   |_|\___/_/\_\\__|
                        //                        |___/                             

                        int[] Sentence_WC = new int[Sentences.Length];
                        VaderSharp.SentimentAnalysisResults[] results = new VaderSharp.SentimentAnalysisResults[Sentences.Length];

                        for (int i = 0; i < Sentences.Length; i++)
                        {
                            results[i] = VADER.PolarityScores(Sentences[i]);
                            Sentence_WC[i] = Sentences[i].Split().Where(x => !string.IsNullOrWhiteSpace(x)).ToArray().Length;
                        }







                        // __        __    _ _          ___        _               _   
                        // \ \      / / __(_) |_ ___   / _ \ _   _| |_ _ __  _   _| |_ 
                        //  \ \ /\ / / '__| | __/ _ \ | | | | | | | __| '_ \| | | | __|
                        //   \ V  V /| |  | | ||  __/ | |_| | |_| | |_| |_) | |_| | |_ 
                        //    \_/\_/ |_|  |_|\__\___|  \___/ \__,_|\__| .__/ \__,_|\__|
                        //                                            |_|              




                        string[] OutputString = new string[8];
                        OutputString[0] = "\"" + Filename_Clean + "\"";
                        OutputString[1] = "0";
                        OutputString[2] = TotalSentences.ToString();
                        OutputString[3] = "";

                        int TotalWC = 0;

                        if (TotalStringLength > 0)
                        {

                            Dictionary<string, double> Average_Results = new Dictionary<string, double>();
                            Average_Results.Add("Positive", 0.0);
                            Average_Results.Add("Neutral", 0.0);
                            Average_Results.Add("Negative", 0.0);
                            Average_Results.Add("Compound", 0.0);

                            for (int i = 0; i < TotalSentences; i++)
                            {
                                TotalWC += Sentence_WC[i];
                                Average_Results["Positive"] += results[i].Positive;
                                Average_Results["Neutral"] += results[i].Neutral;
                                Average_Results["Negative"] += results[i].Negative;
                                Average_Results["Compound"] += results[i].Compound;

                                //write the sentence-level output
                                string[] OutputString_Sentence_Level = new string[8];
                                OutputString_Sentence_Level[0] = "\"" + Filename_Clean + "\"";
                                OutputString_Sentence_Level[1] = Sentence_WC[i].ToString();
                                OutputString_Sentence_Level[2] = "\"" + Sentences[i].Replace("\"", "\"\"") + "\"";
                                OutputString_Sentence_Level[3] = "";

                                if (results[i].Compound > 0.05)
                                {
                                    OutputString_Sentence_Level[3] = "pos";
                                }
                                else if (results[i].Compound > -0.05)
                                {
                                    OutputString_Sentence_Level[3] = "neut";
                                }
                                else
                                {
                                    OutputString_Sentence_Level[3] = "neg";
                                }

                                OutputString_Sentence_Level[4] = results[i].Compound.ToString();
                                OutputString_Sentence_Level[5] = results[i].Positive.ToString();
                                OutputString_Sentence_Level[6] = results[i].Negative.ToString();
                                OutputString_Sentence_Level[7] = results[i].Neutral.ToString();

                                outputFileSentences.WriteLine(String.Join(",", OutputString_Sentence_Level));



                            }

                            Average_Results["Positive"] = Average_Results["Positive"] / (double)TotalSentences;
                            Average_Results["Neutral"] = Average_Results["Neutral"] / (double)TotalSentences;
                            Average_Results["Negative"] = Average_Results["Negative"] / (double)TotalSentences;
                            Average_Results["Compound"] = Average_Results["Compound"] / (double)TotalSentences;


                            OutputString[1] = TotalWC.ToString();
                            OutputString[4] = Average_Results["Compound"].ToString();
                            OutputString[5] = Average_Results["Positive"].ToString();
                            OutputString[6] = Average_Results["Negative"].ToString();
                            OutputString[7] = Average_Results["Neutral"].ToString();

                            if (Average_Results["Compound"] > 0.05)
                            {
                                OutputString[3] = "pos";
                            }
                            else if (Average_Results["Compound"] > -0.05)
                            {
                                OutputString[3] = "neut";
                            }
                            else
                            {
                                OutputString[3] = "neg";
                            }


                        }


                        else
                        {
                            OutputString[2] = "";
                            for (int i = 3; i < 8; i++) OutputString[i + 3] = "";
                        }


                        outputFile.WriteLine(String.Join(",", OutputString));








                    }


                }

            }

            }
            catch
            {
                MessageBox.Show("VADER-Tots encountered an issue somewhere while trying to analyze your texts. The most common cause of this is trying to open your output file while VADER-Tots is still running. Did any of your input files move, or is your output file being opened/modified by another application?", "Error while analyzing", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



        }


        //when the bgworker is done running, we want to re-enable user controls and let them know that it's finished
        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            StartButton.Enabled = true;
            ScanSubfolderCheckbox.Enabled = true;
            EncodingDropdown.Enabled = true;
            FilenameLabel.Text = "Finished!";
            MessageBox.Show("VADER-Tots has finished analyzing your texts.", "Analysis Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }







        public class DictionaryData
        {

            public string TextFileFolder { get; set; }
            public string OutputFileLocation { get; set; }
            
        }

        //https://stackoverflow.com/a/24367618
        string AddSuffix(string filename, string suffix)
        {
            string fDir = Path.GetDirectoryName(filename);
            string fName = Path.GetFileNameWithoutExtension(filename);
            string fExt = Path.GetExtension(filename);
            return Path.Combine(fDir, String.Concat(fName, suffix, fExt));
        }















    }



}
