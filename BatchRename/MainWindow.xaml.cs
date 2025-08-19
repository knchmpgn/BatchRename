using BatchRename;
using Microsoft.Win32;
using BatchRename.Classes;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace BatchRename
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string githubLink = "https://github.com/Masaz-/BatchRename";

        public ObservableCollection<BatchFile> Files { get; set; }
        public ObservableCollection<BatchRule> Rules { get; set; }

        readonly Random rnd = new Random();
        RuleWindow ruleWindow = null;

        public MainWindow()
        {
            InitializeComponent();

            Files = new ObservableCollection<BatchFile>();
            Rules = new ObservableCollection<BatchRule>();

            Rules.CollectionChanged += Rules_CollectionChanged;

            DataContext = this;

            // Process command line arguments on startup
            ProcessCommandLineArguments();
        }

        /// <summary>
        /// Process command line arguments to add files passed from context menu
        /// </summary>
        private void ProcessCommandLineArguments()
        {
            try
            {
                string[] args = Environment.GetCommandLineArgs();
                
                // Skip the first argument (executable path)
                if (args.Length > 1)
                {
                    List<BatchFile> tmpFiles = new List<BatchFile>();
                    
                    for (int i = 1; i < args.Length; i++)
                    {
                        string filePath = args[i];
                        
                        // Validate file exists
                        if (File.Exists(filePath))
                        {
                            tmpFiles.Add(new BatchFile(filePath));
                        }
                        else if (Directory.Exists(filePath))
                        {
                            // If directory is passed, add all files in directory
                            string[] directoryFiles = Directory.GetFiles(filePath);
                            foreach (string file in directoryFiles)
                            {
                                tmpFiles.Add(new BatchFile(file));
                            }
                        }
                    }
                    
                    // Sort files by name
                    tmpFiles = tmpFiles.OrderBy(f => f.Name).ToList();
                    
                    // Add files to collection
                    foreach (BatchFile file in tmpFiles)
                    {
                        Files.Add(file);
                    }
                    
                    // Apply rules if any files were added
                    if (tmpFiles.Count > 0)
                    {
                        ApplyRules();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing command line arguments: {ex.Message}", 
                    "Command Line Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Public method to add files programmatically (for external integration)
        /// </summary>
        public void AddFiles(string[] filePaths)
        {
            try
            {
                List<BatchFile> tmpFiles = new List<BatchFile>();
                
                foreach (string filePath in filePaths)
                {
                    if (File.Exists(filePath))
                    {
                        tmpFiles.Add(new BatchFile(filePath));
                    }
                }
                
                tmpFiles = tmpFiles.OrderBy(f => f.Name).ToList();
                
                foreach (BatchFile file in tmpFiles)
                {
                    Files.Add(file);
                }
                
                if (tmpFiles.Count > 0)
                {
                    ApplyRules();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding files: {ex.Message}", 
                    "Add Files Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        long LongRandom(long min, long max, Random rand)
        {
            byte[] buf = new byte[8];
            rand.NextBytes(buf);
            long longRand = BitConverter.ToInt64(buf, 0);

            return (Math.Abs(longRand % (max - min)) + min);
        }

        private void OpenRuleWindow(BatchRule rule = null)
        {
            ruleWindow = new RuleWindow(rule);
            ruleWindow.Closed += RuleWindow_Closed;
            ruleWindow.ShowDialog();
        }

        private string[] OpenFiles(string path, string filter, out string pickedPath)
        {
            string[] files;
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = filter,
                InitialDirectory = path
            };
            var result = dialog.ShowDialog();

            if (result.HasValue && dialog.FileNames.Length > 0)
            {
                pickedPath = Path.GetDirectoryName(dialog.FileNames[0]);
                files = dialog.FileNames;
            }
            else
            {
                pickedPath = "";
                files = new string[0];
            }

            return files;
        }

        private void RenameFiles()
        {
            for (int i = 0; i < Files.Count; i++)
            {
                try
                {
                    Files[i].Rename();
                }
                catch (Exception ex)
                {
                    Files[i].Status = "Could not rename the file: " + ex.Message;
                }
            }
        }

        private void PickFiles(string path, out string pickedPath)
        {
            List<BatchFile> tmpFiles = new List<BatchFile>();
            string[] filenames = OpenFiles(path, "All Files|*.*", out pickedPath);

            foreach (string filename in filenames)
            {
                tmpFiles.Add(new BatchFile(filename));
            }

            tmpFiles = new List<BatchFile>(tmpFiles.OrderByDescending(f => f.Name));

            foreach (BatchFile file in tmpFiles)
            {
                Files.Add(file);
            }

            ApplyRules();
        }

        private void SaveRules()
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                DefaultExt = ".xml",
                Filter = "BatchRule files (.xml)|*.xml"
            };

            var result = dlg.ShowDialog();

            if (result == true)
            {
                List<BatchRule> listOfRules = new List<BatchRule>();

                foreach (BatchRule rule in Rules)
                {
                    listOfRules.Add(rule);
                }

                BatchSerializer.WriteToXmlFile(dlg.FileName, listOfRules);
            }
        }

        private void ApplyRules()
        {
            try
            {
                TextInfo ti = new CultureInfo("fi-Fi", false).TextInfo;

                for (int f = 0; f < Files.Count; f++)
                {
                    string NewName = Files[f].Name;
                    string ext = Path.GetExtension(NewName);

                    for (int r = 0; r < Rules.Count; r++)
                    {
                        string extLessName = Path.GetFileNameWithoutExtension(NewName);
                        string fullname = Rules[r].Extension ? extLessName + ext : extLessName;

                        // Apply rules

                        if (Rules[r].Insert) // Insert text
                        {
                            if (Rules[r].InsertTextAt == -1 || Rules[r].InsertTextAt > fullname.Length)
                            {
                                fullname += Rules[r].InsertText;
                            }
                            else
                            {
                                fullname = fullname.Insert(Rules[r].InsertTextAt, Rules[r].InsertText);
                            }
                        }

                        if (Rules[r].Replace && Rules[r].ReplaceText != "") // Replace text
                        {
                            if (Rules[r].ReplaceTextWith == null)
                            {
                                Rules[r].ReplaceTextWith = "";
                            }

                            if (Rules[r].ReplaceIsRegex)
                            {
                                Regex regex = new Regex(Rules[r].ReplaceText);

                                fullname = regex.Replace(fullname, Rules[r].ReplaceTextWith);
                            }
                            else
                            {
                                fullname = fullname.Replace(Rules[r].ReplaceText, Rules[r].ReplaceTextWith);
                            }
                        }

                        if (Rules[r].Remove && Rules[r].RemoveStartText != "" && Rules[r].RemoveEndText != "") // Remove text
                        {
                            int start = 0;
                            int end = fullname.Length;

                            if (Rules[r].RemoveStartIsNumber)
                            {
                                int.TryParse(Rules[r].RemoveStartText, out start);

                                if (start == -1)
                                {
                                    start = fullname.Length;
                                }
                            }
                            else
                            {
                                start = fullname.IndexOf(Rules[r].RemoveStartText);
                            }

                            if (start > fullname.Length)
                            {
                                start = fullname.Length;
                            }

                            if (Rules[r].RemoveEndIsNumber)
                            {
                                int.TryParse(Rules[r].RemoveEndText, out end);

                                if (end == -1)
                                {
                                    end = fullname.Length;
                                }
                            }
                            else
                            {
                                if (Rules[r].RemoveEndText == "")
                                {
                                    end = fullname.IndexOf(Rules[r].RemoveEndText);
                                }
                                else
                                {
                                    end = fullname.Length;
                                }
                            }

                            if (end > fullname.Length)
                            {
                                end = fullname.Length;
                            }

                            if (start == end)
                            {
                                fullname = fullname.Remove(start);
                            }
                            else
                            {
                                if (start == -1)
                                {
                                    start = 0;
                                }

                                if (end == -1)
                                {
                                    end = fullname.Length;
                                }

                                fullname = fullname.Remove(start, end - start);
                            }
                        }

                        if (Rules[r].CasingLowercase) // Convert casing to lowercase
                        {
                            fullname = fullname.ToLower();
                        }

                        if (Rules[r].CasingUppercase) // Convert casing to uppercase
                        {
                            fullname = fullname.ToUpper();
                        }

                        if (Rules[r].CasingUppercaseWords) // Uppercase first letter of each word
                        {
                            fullname = ti.ToTitleCase(fullname);
                        }

                        if (Rules[r].CasingUppercaseFirstWord) // Uppercase first letter of text
                        {
                            if (fullname.Length > 0)
                            {
                                fullname = char.ToUpper(fullname[0]) + fullname.Substring(1);
                            }
                        }

                        if (Rules[r].TrimWhitespace) // Trim whitespace
                        {
                            fullname = fullname.Trim();
                        }

                        if (Rules[r].CleanDoubleSpaces) // Clean double space
                        {
                            fullname = fullname.Replace("   ", " ").Replace("  ", " ");
                        }

                        if (Rules[r].RandomNumbering) // Add random numbering
                        {
                            long rndMin = (long)Math.Pow(10, Rules[r].RandomNumbersCount);
                            long rndMax = rndMin * 10;

                            if (Rules[r].RandomNumbersAt == -1 || Rules[r].RandomNumbersAt > fullname.Length)
                            {
                                fullname += LongRandom(rndMin, rndMax, rnd);
                            }
                            else
                            {
                                fullname = fullname.Insert(Rules[r].RandomNumbersAt, LongRandom(rndMin, rndMax, rnd).ToString());
                            }
                        }

                        if (Rules[r].RandomizeFilenames) // Randomize filenames
                        {
                            long rndMin = 100000000;
                            long rndMax = rndMin * 10;

                            fullname = LongRandom(rndMin, rndMax, rnd).ToString();
                        }

                        NewName = fullname;

                        if (!Rules[r].Extension) // Add extension back
                        {
                            NewName += ext;
                        }
                    }

                    Files[f].NewName = NewName;
                }
            }
            catch (Exception ex)
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");

                MessageBox.Show(ex.Message + ex.StackTrace, "Applying Rules Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Rules

        private void Rules_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            for (var i = 0; i < Rules.Count; i++)
            {
                Rules[i].Name = "Rule " + (i+1);
            }

            ApplyRules();
        }

        private void BtnRemoveRule_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = DgRules.SelectedIndex;

            if (selectedIndex != -1)
            {
                Rules.RemoveAt(selectedIndex);

                if (Rules.Count > 0)
                {
                    if (Rules.Count - 1 < selectedIndex)
                    {
                        selectedIndex--;
                    }

                    DgRules.SelectedIndex = selectedIndex;
                }
            }

            ApplyRules();
        }

        private void BtnUp_Click(object sender, RoutedEventArgs e)
        {
            if (DgRules.SelectedIndex != -1 && DgRules.SelectedIndex > 0)
            {
                Rules.Move(DgRules.SelectedIndex, DgRules.SelectedIndex - 1);
            }

            ApplyRules();
        }

        private void BtnDown_Click(object sender, RoutedEventArgs e)
        {
            if (DgRules.SelectedIndex != -1 && DgRules.SelectedIndex < Rules.Count)
            {
                Rules.Move(DgRules.SelectedIndex, DgRules.SelectedIndex + 1);
            }

            ApplyRules();
        }

        private void BtnAddRule_Click(object sender, RoutedEventArgs e)
        {
            OpenRuleWindow();
        }

        private void BtnAddRuleFromFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string[] filepaths = OpenFiles(Properties.Settings.Default.LastLocation, "BatchRule Files (.xml)|*.xml", out string pickedPath);

                Properties.Settings.Default.LastLocation = pickedPath;
                Properties.Settings.Default.Save();

                foreach (string filepath in filepaths)
                {
                    List<BatchRule> r = BatchSerializer.ReadFromXmlFile<List<BatchRule>>(filepath);

                    foreach (BatchRule rule in r)
                    {
                        Rules.Add(rule);
                    }
                }
            }
            catch (Exception ex)
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");

                MessageBox.Show("Failed to read from file: " + ex.Message + ex.StackTrace, "Reading Rules from a File Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnSaveRuleToFile_Click(object sender, RoutedEventArgs e)
        {
            if (Rules.Count > 0)
            {
                SaveRules();
            }
        }

        private void DgRules_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OpenRuleWindow((BatchRule)DgRules.SelectedItem);
        }

        private void DgRules_MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.OriginalSource is ScrollViewer)
            {
                if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                {
                    ((DataGrid)sender).UnselectAll();
                }
                else if (e.ChangedButton == System.Windows.Input.MouseButton.Right)
                {
                    OpenRuleWindow();
                }
            }
        }

        // Files

        private void DgFiles_Drop(object sender, DragEventArgs e)
        {
            try
            {
                List<BatchFile> tmpFiles = new List<BatchFile>();

                foreach (string filename in (string[])e.Data.GetData(DataFormats.FileDrop))
                {
                    tmpFiles.Add(new BatchFile(filename));
                }

                tmpFiles = new List<BatchFile>(tmpFiles.OrderBy(f => f.Name));

                foreach (BatchFile file in tmpFiles)
                {
                    Files.Add(file);
                }

                ApplyRules();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace, "Dropping Files Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DgFiles_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            PickFiles("", out string pickedPath);
        }

        private void BtnSortUp_Click(object sender, RoutedEventArgs e)
        {
            Files = new ObservableCollection<BatchFile>(Files.OrderBy(f => f.Name));

            DgFiles.ItemsSource = Files;

            ApplyRules();
        }

        private void BtnSortDown_Click(object sender, RoutedEventArgs e)
        {
            Files = new ObservableCollection<BatchFile>(Files.OrderByDescending(f => f.Name));

            DgFiles.ItemsSource = Files;

            ApplyRules();
        }

        private void BtnRemoveFile_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = DgFiles.SelectedIndex;

            if (selectedIndex != -1)
            {
                Files.RemoveAt(selectedIndex);

                if (Files.Count > 0)
                {
                    if (Files.Count - 1 < selectedIndex)
                    {
                        selectedIndex--;
                    }

                    DgFiles.SelectedIndex = selectedIndex;
                }
            }
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            PickFiles("", out string pickedPath);
        }

        // Bottom

        private void BtnClearRules_Click(object sender, RoutedEventArgs e)
        {
            Rules.Clear();
        }

        private void BtnClearFiles_Click(object sender, RoutedEventArgs e)
        {
            Files.Clear();
        }

        private void BtnRename_Click(object sender, RoutedEventArgs e)
        {
            RenameFiles();
        }

        private void RuleWindow_Closed(object sender, EventArgs e)
        {
            RuleWindow rw = (RuleWindow)sender;

            if (rw.Save)
            {
                if (rw.Editing)
                {
                    for (int i = 0; i < Rules.Count; i++)
                    {
                        if (Rules[i].Id == rw.Rule.Id)
                        {
                            Rules[i] = rw.Rule;
                            i = Rules.Count;
                        }
                    }
                }
                else
                {
                    Rules.Add(rw.Rule);
                }
            }
        }

        private void LinkGithub_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo(githubLink));
        }

        private void DgRules_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void DgFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}