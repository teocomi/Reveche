using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO.Packaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Input;

namespace Reveche
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    private const string StreamName = "BasicFileInfo";

    private static Regex FoundYear = new Regex(@"\s\d{4}\s");
    private static Regex FoundYear2019 = new Regex(@"^(\d{4})\u0012");

    private ObservableCollection<RevitFile> SourceCollection = new ObservableCollection<RevitFile>();
    List<String> RevitFiles = new List<string>();
    
    public MainWindow()
    {
      InitializeComponent();
      Title = "Revit Version Checker v" + Assembly.GetExecutingAssembly().GetName().Version;
      SourceFilesList.ItemsSource = SourceCollection;
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

          }

          else if (File.Exists(s) &&
                   (s.ToLower().EndsWith(".rvt") || s.ToLower().EndsWith(".rfa")))
          {
            RevitFiles.Add(s);
          }
            
        }
        //MessageBox.Show("Completed in " + (DateTime.Now - start).TotalSeconds.ToString() + " seconds.");
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

    private static void GetYearFromFileInfo(string pathToRevitFile, RevitFile rf)
    {
      try
      {
        var rawData = GetRawBasicFileInfo(
            pathToRevitFile);

        var rawString = System.Text.Encoding.Unicode
            .GetString(rawData);

        var fileInfoData = rawString.Split(
            new string[] { "\0", "\r\n" },
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
              break;
            }

          }

        }

        //2019+ don't use the conventions above :(
        if (string.IsNullOrEmpty(rf.Version))
        {
          foreach (var info in fileInfoData.Select((val, i) => new { i, val }))
          {
            Match yearMatch = FoundYear2019.Match(info.val);
          
            if (yearMatch.Success) 
            {
              rf.Version = yearMatch.Groups[1].Value;
              rf.AdditionalInfo = "Build: " + fileInfoData[info.i+1];
            }
          }
        }
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
      }
    }
   
    /// <summary>
    /// Dumb way to get current RenameType
    /// </summary>
    /// <returns></returns>
    private Action GetSelectedAction()
    {
      if (PrependRb.IsChecked.Value)
        return Action.Prepend;
      if (AppendRb.IsChecked.Value)
        return Action.Append;
      return Action.Clear;
    }

    #region events

    /// <summary>
    /// Click Remove Files
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void RemoveFilesBtn_OnClick(object sender, RoutedEventArgs e)
    {
      try
      {
        var selectetitems = SourceFilesList.SelectedItems.Cast<RevitFile>().ToList();
        foreach (var clashItem in selectetitems)
          SourceCollection.Remove(clashItem);
      }
      catch (System.Exception ex1)
      {
        MessageBox.Show("exception: " + ex1);
      }
    }

    /// <summary>
    /// Click Add Files
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void AddFilesBtn_OnClick(object sender, RoutedEventArgs e)
    {
      try
      {
        var openFileDialog1 = new Microsoft.Win32.OpenFileDialog();
        openFileDialog1.Title = "Select files";
        openFileDialog1.Filter = "Revit Files (*.rvt;*.rfa)|*.rvt;*.rfa|All files (*.*)|*.*";
        openFileDialog1.Multiselect = true;

        openFileDialog1.RestoreDirectory = true;
        var result = openFileDialog1.ShowDialog(); // Show the dialog.

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
    /// <summary>
    /// Click Add Files from List
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void AddFilesFromListBtn_OnClick( object sender, RoutedEventArgs e ) {
      try {
        var openFileDialog1 = new Microsoft.Win32.OpenFileDialog();
        openFileDialog1.Title = "Select file list";
        openFileDialog1.Filter = "Text Files (*.txt)|*.txt|All files (*.*)|*.*";
        openFileDialog1.Multiselect = false;

        openFileDialog1.RestoreDirectory = true;
        var result = openFileDialog1.ShowDialog(); // Show the dialog.

        if ( result == true ) // Test result.
        {
          // Parse file to extract valid filenames into a array equivalent to the open file dialog.

          var fileList = new List<string>();

          if ( File.Exists( openFileDialog1.FileName ) ) {

            System.IO.StreamReader file = new System.IO.StreamReader( openFileDialog1.FileName );
            string line;
            while ( ( line = file.ReadLine() ) != null ) {
            
              if ( File.Exists( line ) ) {
                fileList.Add( line );
              }
            }

            ProcessSourceFiles( fileList.ToArray() );
          }

        }
      }
      catch ( System.Exception ex1 ) {
        MessageBox.Show( "exception: " + ex1 );
      }

    }


    /// <summary>
    /// Update names after toggle has changed
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
    {
      try
      {
        foreach (var revitFile in SourceCollection)
        {
          revitFile.RefreshName(GetSelectedAction());
        }
      }
      catch (System.Exception ex1)
      {
        MessageBox.Show("exception: " + ex1);
      }
    }

    /// <summary>
    /// Start renming process
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
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
        if (i != 0)
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
    
    private void Window_Paste( object sender, ExecutedRoutedEventArgs e ) {

      try {
        string text = Clipboard.GetText() as string;

        if ( string.IsNullOrWhiteSpace( text ) )
          return;

        string[] lines = Regex.Split( text, "\r\n|\r|\n" );

        List<string> fileList = new List<string>();

        foreach ( var line in lines ) {

          if ( File.Exists( line ) ) {
            fileList.Add( line );
          }
        }

        ProcessSourceFiles( fileList.ToArray() );
      }
      catch ( System.Exception ex1 ) {
        MessageBox.Show( "exception: " + ex1 );
      }
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
          string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

          ProcessSourceFiles(files);
        }
      }
      catch (System.Exception ex1)
      {
        MessageBox.Show("exception: " + ex1);
      }
    }
    #endregion
   

    #region version stuff

    // Code thanks to Jeremy Tammik
    // Please check his blog
    // http://thebuildingcoder.typepad.com/blog/2013/01/basic-file-info-and-rvt-file-version.html

    private static byte[] GetRawBasicFileInfo(
        string revitFileName)
    {
      try
      {
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
              = (StorageInfo)InvokeStorageRootMethod(
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
              = (StorageInfo)InvokeStorageRootMethod(
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

      private static object InvokeStorageRootMethod(StorageInfo storageRoot, string methodName, params object[] methodArgs)
      {
        try
        {
          Type storageRootType = typeof(StorageInfo).Assembly.GetType("System.IO.Packaging.StorageRoot", false, false);

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
