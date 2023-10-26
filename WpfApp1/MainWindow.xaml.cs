using Microsoft.VisualBasic;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string inputFile { get; set; }
        public string outputFolder { get; set; }
        public int variableLength { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            inputFile = "";
            outputFolder = "";
        }


        private void inButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.ValidateNames = false;
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;
            dialog.FileName = "Folder Selection.";

            if (dialog.ShowDialog() == true)
            {
                string folderPath = dialog.FileName;
                inputText.Text = folderPath;
                inputFile = folderPath;
            }
        }
        private void exButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.ValidateNames = false;
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;
            dialog.FileName = "Folder Selection.";

            if (dialog.ShowDialog() == true)
            {
                string folderPath = dialog.FileName;
                outputText.Text = System.IO.Path.GetDirectoryName(folderPath);
                outputFolder = outputText.Text;
            }
        }

        public void writeToLog(string? data)
        {
            if (debug != null)
                debug.Text += $"\n{data}";
        }

        
        public void writeToLog(IEnumerable<string> items)
        {
            if (debug != null)
            {
                debug.Text += $"\n";
                foreach (String item in items)
                {
                    debug.Text += $",{item}";
                }
            }
        }

        public void Process()
        {
            HashSet<string> output = new HashSet<string>();
            //using OfficeOpenXml;
            
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(inputFile)))
            {
                // If you use EPPlus in a noncommercial context
                // according to the Polyform Noncommercial license:
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                var sb = new StringBuilder(); //this is your data
                for (int rowNum = 2; rowNum <= totalRows; rowNum++) //select starting row here
                {
                    //get an entire row (length = totalColumns)
                    var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                    sb.Append(string.Concat(",", string.Join(",", row)));
                }
                //writeToLog(sb.ToString());
                var splitSb = sb.ToString().Split(',');
                output = splitSb.ToHashSet();
            }
            output.Remove("");      //remove empty string
            //writeToLog(output);
            WriteResultFile(output);
            
        }

        private void WriteResultFile(HashSet<String>output)
        {
            String res = "";
            String fileName = "output";
            List<string> tempList = output.ToList<String>();
            int length = tempList.Count();
            int cntr = 0;
            int fileCntr = 0;
            while (cntr < length) 
            {
                int nextCntr = (cntr + variableLength) < length ?
                                            (cntr + variableLength) :
                                            length;
                using (StreamWriter outputFile = new StreamWriter(System.IO.Path.Combine(outputFolder, $"{fileName}_{fileCntr}.lab")))
                {
                    res = "[RAMCELL]\n";
                    for(int i = cntr; i < nextCntr; i++)
                    {
                        res += $"{tempList[i]}\n";
                    }
                    res += "[LABEL]";
                    
                    outputFile.WriteLine(res);
                }

                using (StreamWriter outputFile = new StreamWriter(System.IO.Path.Combine(outputFolder, $"{fileName}_{fileCntr}.txt")))
                {
                    res = "";
                    for (int i = cntr; i < nextCntr; i++)
                    {
                        res += $"Port: {tempList[i]}, ";
                    }
                    
                    outputFile.WriteLine(res);
                }
                cntr = nextCntr;
                fileCntr++;
            }
            
        }

        private void runButton_Click(object sender, RoutedEventArgs e)
        {
            inputFile = inputText.Text;
            outputFolder = outputText.Text;
            if (System.IO.Path.IsPathFullyQualified(inputFile) &&
                System.IO.Path.IsPathFullyQualified(outputFolder)) 
            {
                // put this process to thread.
                variableLength = (int)slider.Value;
                if (variableLength <= 0)
                {
                    debug.Text = "There's nothing to do...";
                }
                else if (!System.IO.File.Exists(inputFile))
                {
                    debug.Text = "File does not exist!";
                }
                else
                {
                    if (!System.IO.Directory.Exists(outputFolder))
                    {
                        System.IO.Directory.CreateDirectory(outputFolder);
                    }
                    Thread mythread = new Thread(new ThreadStart(Process));
                    mythread.Start();
                    debug.Text = "Processing...";
                    mythread.Join();
                    debug.Text = $"File created at {outputFolder}";
                }
            } else
            {
                debug.Text = "Invalid input file path or output file path!";
            }
            
        }

        private void onSliderChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (sliderVal != null)
                sliderVal.Text = slider.Value.ToString();
        }
    }
}
