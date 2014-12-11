//
// (C) Copyright 2014 by Matteo Cominetti
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// I PROVIDE THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// I SPECIFICALLY DISCLAIM ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. 
// I DO NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.





using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Input;

namespace RVTVersioner
{
    class Program
    {
        private const string StreamName = "BasicFileInfo";
        private static Regex EndYear = new Regex(@"_\d{4}$");
        private static Regex FoundYear = new Regex(@"\s\d{4}\s");
        /// <summary>
        /// Main function
        /// </summary>
        /// <param name="args">dragged files and folders</param>
        static void Main(string[] args)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                Console.Write(@"
     ___ __     _________ __     __                                             
    /    ) |   /    /       |   /                                             
---/___ /--|--/----/--------|--/-----__----__---__---'----__----__----__----__-
  /    |   | /    /         | /    /___) /   ) (_ ` /   /   ) /   ) /___) /   )
_/_____|___|/____/__________|/____(___ _/_____(__)_/___(___/_/___/_(___ _/_____

RVT Versioner v." + version + @"
Free Tool produced by Matteo Cominetti
matteocominetti.com


"
                    );
                List<String> RevitFiles = new List<string>();

                //seach subdirectories for revit files
                foreach (var s in args)
                {
                    if (Directory.Exists(s))
                        RevitFiles.AddRange(
                            Directory
                            .EnumerateFiles(s)
                            .Where(file => file.ToLower().EndsWith("rvt") || file.ToLower().EndsWith("rfa") || s.ToLower().EndsWith("rte") || s.ToLower().EndsWith("rtf"))
                            .ToList());
                    else if (File.Exists(s) && (s.ToLower().EndsWith("rvt") || s.ToLower().EndsWith("rfa") || s.ToLower().EndsWith("rte") || s.ToLower().EndsWith("rtf")))
                        RevitFiles.Add(s);

                }
                //user double clicked on the exe
                if (args.Length == 0)
                {
                    Console.WriteLine(@"USAGE:" +
                                      "Drag and Drop one or more files/folders on this executable. " +
                                      "You will be prompted if you want to rename the Revit files found" +
                                      " appending the _YEAR to the filename or if you just want to clean their names.");

                    Console.ReadKey();
                    return;
                }
                //no valid files found
                if (RevitFiles.Count == 0)
                {
                    Console.WriteLine("No Revit files found!");
                    Console.ReadKey();
                    return;
                }

                ProcessFiles(RevitFiles, ConsoleKey.H);
                String plural = RevitFiles.Count == 1 ? "" : "s";
                Console.WriteLine(@"
I found " + RevitFiles.Count + " Revit file" + plural + ", do you want to Rename, cLean or Cancel?" +
                    @"
Press one of the following keys on your keyboard to continue:
R = Rename all Revit files appending _YEAR
L = cLean, removes _YEAR from all .rvt files that have it
C = Cancel and quit application");
                while (true)
                {
                    //Program.DisplayMessage("The file: " + binaryFileName + " already exist. Do you want to overwrite it? Y/N");
                    ConsoleKeyInfo answer = Console.ReadKey();

                    if (answer.Key.Equals(ConsoleKey.C))
                        return;
                    if (answer.Key.Equals(ConsoleKey.R) || answer.Key.Equals(ConsoleKey.L))
                    {
                        ProcessFiles(RevitFiles, answer.Key);
                        Console.ReadKey();
                        break;
                    }
                    Console.WriteLine();
                    Console.WriteLine("Invalid Key pressed. Please try again.");

                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }

        private static void ProcessFiles(List<String> RevitFiles, ConsoleKey key)
        {
            Console.WriteLine();
            if (key.Equals(ConsoleKey.H))
                Console.WriteLine("{0,-4} {1,-4} {2,-30}",
                    "YEAR",
                    "EXT",
                    "FILENAME");
            else
                Console.WriteLine("{0,-4} {1,-4} {2,-30} {3,-30}",
                           "YEAR",
                           "EXT",
                           "FILENAME",
                           "NEWNAME");

            foreach (var revitFile in RevitFiles)
            {
                if (!StructuredStorageUtils.IsFileStucturedStorage(
              revitFile) || !File.Exists(revitFile))
                {
                    Console.WriteLine("*** " + Path.GetFileName(revitFile) + " is not valid!");
                    continue;
                }
                    ProcessFilesLoop(revitFile, key);
            }
           
        }

        private static void ProcessFilesLoop(string pathToRevitFile, ConsoleKey key)
        {
            try
            {

                string version = "????";
                char[] charsToTrim = { ' ' };
                string filename = Path.GetFileNameWithoutExtension(pathToRevitFile);
                string newname = "";
                string extension = Path.GetExtension(pathToRevitFile);
                string cleanfilename = EndYear.Replace(filename, "");
                //remove trailing spaces
                cleanfilename = cleanfilename.TrimEnd(charsToTrim);
                string directory = Path.GetDirectoryName(pathToRevitFile);
                
                if (key.Equals(ConsoleKey.L))
                {

                    Match match = EndYear.Match(filename);
                    if (match.Success)
                    {
                        //version is the appended year to the file
                        version = match.Value.Replace("_", "");
                        newname = cleanfilename + extension;
                        File.Move(pathToRevitFile,
                            Path.Combine(directory, newname));
                        
                    }
                    else
                        return;
                }
                else
                {
   
                    //version is the year found in the file info
                    version = GetYearFromFileInfo(pathToRevitFile);
                    if (key.Equals(ConsoleKey.R) && version != "????")
                    {
                        //rename file
                        newname = cleanfilename + "_" + version + extension;
                        File.Move(pathToRevitFile,
                            Path.Combine(directory, newname));
                    }
                }

                Console.WriteLine("{0,-4} {1,-4} {2,-30} {3,-30}",
                    version, extension, filename, newname);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }

        private static string GetYearFromFileInfo(string pathToRevitFile)
        {
            string version = "????";
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
                            return match.Value.Replace(" ", "");
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            return version;
        }
  

        // Code thanks to Jeremy Tammik
        // Please check his blog
        // http://thebuildingcoder.typepad.com/blog/2013/01/basic-file-info-and-rvt-file-version.html

        private static byte[] GetRawBasicFileInfo(
          string revitFileName)
        {

            if (!StructuredStorageUtils.IsFileStucturedStorage(
              revitFileName))
            {
                Console.WriteLine(
                  "File is not a structured storage file");
                Console.ReadKey();
                throw new NotSupportedException("File is not a structured storage file");
            }

            using (StructuredStorageRoot ssRoot =
                new StructuredStorageRoot(revitFileName))
            {
                if (!ssRoot.BaseRoot.StreamExists(StreamName))
                {
                    Console.WriteLine((string.Format(
                      "File doesn't contain {0} stream", StreamName)));
                    Console.ReadKey();
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

    }

    public static class StructuredStorageUtils
    {
        [DllImport("ole32.dll")]
        static extern int StgIsStorageFile(
          [MarshalAs(UnmanagedType.LPWStr)]
      string pwcsName);

        public static bool IsFileStucturedStorage(
          string fileName)
        {
            int res = StgIsStorageFile(fileName);

            if (res == 0)
                return true;

            if (res == 1)
                return false;
            Console.WriteLine("File not found");
            Console.ReadKey();
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
        StorageInfo _storageRoot;

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
                Console.WriteLine("Cannot get StructuredStorageRoot: " + ex.Message);
                Console.ReadKey();
                throw new StructuredStorageException(
                  "Cannot get StructuredStorageRoot", ex);
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
                Console.WriteLine("Cannot get StructuredStorageRoot: " + ex.Message);
                Console.ReadKey();
                throw new StructuredStorageException(
                  "Cannot get StructuredStorageRoot", ex);
            }
        }

        private static object InvokeStorageRootMethod(
          StorageInfo storageRoot,
          string methodName,
          params object[] methodArgs)
        {
            Type storageRootType
              = typeof(StorageInfo).Assembly.GetType(
                "System.IO.Packaging.StorageRoot",
                true, false);

            object result = storageRootType.InvokeMember(
              methodName,
              BindingFlags.Static | BindingFlags.Instance
              | BindingFlags.Public | BindingFlags.NonPublic
              | BindingFlags.InvokeMethod,
              null, storageRoot, methodArgs);

            return result;
        }

        private void CloseStorageRoot()
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

        //private void ControlledApplication_DocumentOpening( 
        //  object sender,
        //  DocumentOpeningEventArgs e )
        //{
        //  FileInfo revitFileToUpgrade 
        //    = new FileInfo( e.PathName );

        //  Regex buildInfoRegex = new Regex(
        //    @"Revit\sArchitecture\s(?<Year>\d{4})\s"
        //    +@"\(Build:\s(?<Build>\w*)\((<Processor>\w{3})\)\)" );

        //  using( StreamReader streamReader =
        //    new StreamReader( e.PathName, Encoding.Unicode ) )
        //  {
        //    string fileContents = streamReader.ReadToEnd();

        //    Match buildInfo = buildInfoRegex.Match( fileContents );
        //    string year = buildInfo.Groups["Year"].Value;
        //    string build = buildInfo.Groups["Build"].Value;
        //    string processor = buildInfo.Groups["Processor"].Value;
        //  }
        //}

    }
}
