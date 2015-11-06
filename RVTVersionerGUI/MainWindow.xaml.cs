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
using System.IO;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Input;

namespace RVTVersionerGUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string StreamName = "BasicFileInfo";

        private static Regex FoundYear = new Regex(@"\s\d{4}\s");
        private ObservableCollection<RevitFile> SourceCollection = new ObservableCollection<RevitFile>();
        List<String> RevitFiles = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            Title = "Revit Version Checker v" + Assembly.GetExecutingAssembly().GetName().Version;
            SourceFilesList.ItemsSource = SourceCollection;
        }


        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            whitespace.Visibility = System.Windows.Visibility.Visible;

        }

        private void Window_DragLeave(object sender, DragEventArgs e)
        {
            whitespace.Visibility = System.Windows.Visibility.Hidden;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            try
            {
                whitespace.Visibility = System.Windows.Visibility.Hidden;
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    // Note that you can have more than one file.
                    string[] files = (string[]) e.Data.GetData(DataFormats.FileDrop);

                    ProcessSourceFiles(files);
                }
            }
            catch (System.Exception ex1)
            {
                MessageBox.Show("exception: " + ex1);
            }
        }

        private void ProcessSourceFiles(string[] files)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;

                RevitFiles = new List<string>();
              var start = DateTime.Now;

                //seach subdirectories for revit files
                foreach (var s in files)
                {
                  if (Directory.Exists(s))
                  {
                    RevitFiles.AddRange(Directory.GetFiles(s, "*.rvt", SearchOption.AllDirectories).ToList());
                    RevitFiles.AddRange(Directory.GetFiles(s, "*.rfa", SearchOption.AllDirectories).ToList());
                    
                    //WalkDirectoryTree(new DirectoryInfo(s));
                  }
                        
                    else if (File.Exists(s) &&
                             (s.ToLower().EndsWith(".rvt") || s.ToLower().EndsWith(".rfa")))
                        RevitFiles.Add(s);

                }
                MessageBox.Show("Completed in "+(DateTime.Now-start).TotalSeconds.ToString() + " seconds.");
                //user double clicked on the exe
                if (files.Length == 0)
                {
                    MessageBox.Show("No files Found!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;

                }
                //no valid files found
                if (RevitFiles.Count == 0)
                {
                    MessageBox.Show("No Revit files found!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // SourceCollection.Clear();
                foreach (var f in RevitFiles)
                {
                    var rf = new RevitFile(f);
                   GetYearFromFileInfo(f, rf);
                    rf.RefreshName(GetSelectedAction());
                    SourceCollection.Add(rf);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
     private void WalkDirectoryTree(System.IO.DirectoryInfo root)
    {
        System.IO.FileInfo[] files = null;
        System.IO.DirectoryInfo[] subDirs = null;

        // First, process all the files directly under this folder 
        try
        {
          var dir = root.FullName;
            RevitFiles.AddRange(
                            Directory
                                .EnumerateFiles(dir)
                                .Where(
                                    file =>
                                        file.ToLower().EndsWith("rvt") || file.ToLower().EndsWith("rfa"))
                                .ToList());
        }
        // This is thrown if even one of the files requires permissions greater 
        // than the application provides. 
        catch (UnauthorizedAccessException e)
        {
        }

        catch (System.IO.DirectoryNotFoundException e)
        {
            Console.WriteLine(e.Message);
        }
        subDirs = root.GetDirectories();
        if (subDirs != null)
        {
            // Now find all the subdirectories under this directory.
            subDirs = root.GetDirectories();

            foreach (System.IO.DirectoryInfo dirInfo in subDirs)
            {
                // Resursive call for each subdirectory.
                WalkDirectoryTree(dirInfo);
            }
        }
    }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            try
            {
                //bool dropEnabled = true;

                //if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
                //{
                //    string[] filenames = e.Data.GetData(DataFormats.FileDrop, true) as string[];

                //    foreach (string filename in filenames)
                //    {
                //        if (Path.GetExtension(filename).ToUpperInvariant() != ".BCFZIP")
                //        {
                //            dropEnabled = false;
                //            break;
                //        }
                //    }

                //}
                //else
                //{
                //    dropEnabled = false;
                //}


                //if (!dropEnabled)
                //{
                //    e.Effects = DragDropEffects.None;
                //    e.Handled = true;
                //}
            }
            catch (System.Exception ex1)
            {
                MessageBox.Show("exception: " + ex1);
            }
        }

        private static string GetYearFromFileInfo(string pathToRevitFile, RevitFile rf)
        {
            string version = "????";
            try
            {
                var rawData = GetRawBasicFileInfo(
                    pathToRevitFile);

                var rawString = System.Text.Encoding.Unicode
                    .GetString(rawData);

                var fileInfoData = rawString.Split(
                    new string[] {"\0", "\r\n"},
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (var info in fileInfoData)
                {
                    //Console.WriteLine(info);
                    if (info.Contains("Autodesk"))
                    {
                        Match match = FoundYear.Match(info);
                        if (match.Success)
                        {
                            rf.Version = match.Value.Replace(" ", "");
                            rf.AdditionalInfo = info.Replace("Revit Build: ", "");
                        }
                             
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return version;
        }

        private void RemoveFilesBtn_OnClick(object sender, RoutedEventArgs e)
        {
            try{

            var selectetitems = SourceFilesList.SelectedItems.Cast<RevitFile>().ToList();
            foreach (var clashItem in selectetitems)
                SourceCollection.Remove(clashItem);
            }
            catch (System.Exception ex1)
            {
                MessageBox.Show("exception: " + ex1);
            }
        }

        private void AddFilesBtn_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog openFileDialog1 = new Microsoft.Win32.OpenFileDialog();
                openFileDialog1.Title = "Select files";
                openFileDialog1.Filter = "Revit Files (*.rvt;*.rfa)|*.rvt;*.rfa|All files (*.*)|*.*";
                openFileDialog1.Multiselect = true;

                //openFileDialog1.DefaultExt = ".bcfzip";
                openFileDialog1.RestoreDirectory = true;
                Nullable<bool> result = openFileDialog1.ShowDialog(); // Show the dialog.

                if (result == true) // Test result.
                {
                    ProcessSourceFiles(openFileDialog1.FileNames);
                }
            }
            catch (System.Exception ex1)
            {
                MessageBox.Show("exception: " + ex1);
            }

        }

        private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            try
            {
            foreach (var revitFile in SourceCollection)
            {
                RadioButton rb = sender as RadioButton;
                revitFile.RefreshName(rb.Content.ToString().ToLower());
            }
             }
            catch (System.Exception ex1)
            {
                MessageBox.Show("exception: " + ex1);
            }
        }

        private void RenameFilesBtn_OnClick(object sender, RoutedEventArgs e)
        {
             try
            {
            if (SourceCollection.GroupBy(n => n.NewFullPath).Any(c => c.Count() > 1))
            {
                MessageBox.Show("Some files have the same output path and name,\ncannot continue!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
                int i = 0;
            foreach (var revitFile in SourceCollection)
            {
                if (File.Exists(revitFile.FullPath) && !string.IsNullOrEmpty(revitFile.NewFilename))
                {
                    if (revitFile.NewFullPath == revitFile.FullPath)
                        continue;
                    
                    if (File.Exists(revitFile.NewFullPath))
                    {
                        i++;
                        continue;
                    }
                    File.Move(revitFile.FullPath, revitFile.NewFullPath);
                }
            }
            if (i!=0)
            {
                MessageBox.Show(i + " file/s already existed and have been skipped,\nplease be careful!", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            }
             catch (System.Exception ex1)
             {
                 MessageBox.Show("exception: " + ex1);
             }
        }

        private string GetSelectedAction ()
        {
            if (PrependRb.IsChecked.Value)
                return "prepend";
            if (AppendRb.IsChecked.Value)
                return "append";
            return "clear";
        }


        #region version stuff


        // Code thanks to Jeremy Tammik
        // Please check his blog
        // http://thebuildingcoder.typepad.com/blog/2013/01/basic-file-info-and-rvt-file-version.html

        private static byte[] GetRawBasicFileInfo(
            string revitFileName)
        {
            try {
            if (!StructuredStorageUtils.IsFileStucturedStorage(
                revitFileName))
            {
                MessageBox.Show(
                    "File is not a structured storage file");
                throw new NotSupportedException("File is not a structured storage file");
            }

            using (StructuredStorageRoot ssRoot =
                new StructuredStorageRoot(revitFileName))
            {
                if (!ssRoot.BaseRoot.StreamExists(StreamName))
                {
                    MessageBox.Show((string.Format(
                        "File doesn't contain {0} stream", StreamName)));
                    throw new NotSupportedException(string.Format(
                        "File doesn't contain {0} stream", StreamName));
                }


                StreamInfo imageStreamInfo =
                    ssRoot.BaseRoot.GetStreamInfo(StreamName);

                using (Stream stream = imageStreamInfo.GetStream(
                    FileMode.Open, FileAccess.Read))
                {
                    byte[] buffer = new byte[stream.Length];
                    stream.Read(buffer, 0, buffer.Length);
                    return buffer;
                }
            }
            }
            catch (System.Exception ex1)
            {
                MessageBox.Show("exception: " + ex1);
            }
            return null;
        }



        public static class StructuredStorageUtils
        {
            [DllImport("ole32.dll")]
            private static extern int StgIsStorageFile(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

            public static bool IsFileStucturedStorage(
                string fileName)
            {
                int res = StgIsStorageFile(fileName);

                if (res == 0)
                    return true;

                if (res == 1)
                    return false;
                MessageBox.Show("File not found");
                throw new FileNotFoundException(
                    "File not found", fileName);
            }
        }

        public class StructuredStorageException : Exception
        {
            public StructuredStorageException()
            {
            }

            public StructuredStorageException(string message)
                : base(message)
            {
            }

            public StructuredStorageException(
                string message,
                Exception innerException)
                : base(message, innerException)
            {
            }
        }

        public class StructuredStorageRoot : IDisposable
        {
            private StorageInfo _storageRoot;

            public StructuredStorageRoot(Stream stream)
            {
                try
                {
                    _storageRoot
                        = (StorageInfo) InvokeStorageRootMethod(
                            null, "CreateOnStream", stream);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Cannot get StructuredStorageRoot: " + ex.Message);
                    //throw new StructuredStorageException(
                    //    "Cannot get StructuredStorageRoot", ex);
                }
            }

            public StructuredStorageRoot(string fileName)
            {
                try
                {
                    _storageRoot
                        = (StorageInfo) InvokeStorageRootMethod(
                            null, "Open", fileName, FileMode.Open,
                            FileAccess.Read, FileShare.Read);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Cannot get StructuredStorageRoot: " + ex.Message);
                    //throw new StructuredStorageException(
                    //    "Cannot get StructuredStorageRoot", ex);
                }
            }

            private static object InvokeStorageRootMethod(StorageInfo storageRoot,string methodName,params object[] methodArgs)
            {
              try
              {
                Type storageRootType = typeof (StorageInfo).Assembly.GetType("System.IO.Packaging.StorageRoot",false, false);
               
                object result = storageRootType.InvokeMember(
                  methodName,
                  BindingFlags.Static | BindingFlags.Instance
                  | BindingFlags.Public | BindingFlags.NonPublic
                  | BindingFlags.InvokeMethod,
                  null, storageRoot, methodArgs);

                return result;
              }
              catch (Exception ex)
              {
              }
              return null;
            }

              private
              void CloseStorageRoot()
            {
                InvokeStorageRootMethod(_storageRoot, "Close");
            }

            #region Implementation of IDisposable

            public void Dispose()
            {
                CloseStorageRoot();
            }

            #endregion

            public StorageInfo BaseRoot
            {
                get { return _storageRoot; }
            }

            #endregion

        }
    }
}
