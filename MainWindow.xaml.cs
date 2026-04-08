using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic.FileIO;

using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Processing;

using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net.Http;

using System.Windows;
// ✅ ImageSharp aliases (fix ambiguity issues)
using ISImage = SixLabors.ImageSharp.Image;
using ISResizeMode = SixLabors.ImageSharp.Processing.ResizeMode;
using ISSize = SixLabors.ImageSharp.Size;

namespace MicroRenamerWPF
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {

    public MainWindow()
    {
      InitializeComponent();
    }

    private int textMode = 0; // 0 = CAPS, 1 = lower, 2 = Title
    string downloadsPath;
    private bool isExpanded = false;
    private double originalHeight;

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      downloadsPath = Path.Combine(
          Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
          "Downloads"
      );


    }

    //The RENAME button
    private void btnRename_Click(object sender, RoutedEventArgs e)
    {
      RenameFiles();
    }
    private void RenameFiles()
    {
      this.Title = "Micro Renamer for Windows - Syntax Communications";

      if (chkRemoveAllText.IsChecked == true && !string.IsNullOrEmpty(txtRemoveText.Text))
      {
        RenameNumbersOnly(downloadsPath);
      }

      if (chkAddDate.IsChecked == true)
      {
        RenameAddDate(downloadsPath);
      }
      if (chkAddText.IsChecked == true)
      {
        RenameAddText(downloadsPath);
      }
      if (chkRemoveSpecial.IsChecked == true)
      {
        renameSpecial(downloadsPath);
      }
      if (chkNumberItems.IsChecked == true)
      {
        RenameNumbers(downloadsPath);
      }
      RenameShortFiles(downloadsPath);
    }

    private void RenameAddDate(string folder)
    {
      string todaysDate = GetTodaysDate();


      var files = Directory.GetFiles(folder, "*", System.IO.SearchOption.AllDirectories);

      foreach (var file in files)
      {
        try
        {
          string directory = Path.GetDirectoryName(file);
          string nameOnly = Path.GetFileNameWithoutExtension(file);
          string extension = Path.GetExtension(file).ToLower();

          // ⛔ Skip unwanted files
          if (extension == ".docx" || extension == ".doc" || extension == ".zip")
            continue;

          // Skip if already has date
          if (nameOnly.StartsWith(todaysDate))
            continue;

          string newName = todaysDate + nameOnly;
          string newPath = Path.Combine(directory, newName + extension);

          if (file == newPath)
            continue;
          File.Move(file, newPath);
        }
        catch { }
      }

    }
    private string GetTodaysDate()
    {
      return string.Concat(DateTime.Now.Month.ToString("D2"), DateTime.Now.ToString("dd"), DateTime.Now.ToString("yy"), "-");
    }
    private void btnUndoDates_Click(object sender, RoutedEventArgs e)
    {
      RemoveDateFromFiles(downloadsPath, GetTodaysDate());
    }
    private void RemoveDateFromFiles(string folderPath, string datePrefix)
    {
      var files = Directory.GetFiles(
      folderPath,
      ".",
      System.IO.SearchOption.AllDirectories
      );

      foreach (var file in files)
      {
        try
        {
          string directory = Path.GetDirectoryName(file);
          string fileName = Path.GetFileNameWithoutExtension(file);
          string extension = Path.GetExtension(file);

          // Only remove if filename starts with today's date
          if (!fileName.StartsWith(datePrefix))
            continue;

          string newName = fileName.Substring(datePrefix.Length);

          string newPath = Path.Combine(directory, newName + extension);

          int counter = 1;
          while (File.Exists(newPath))
          {
            newPath = Path.Combine(directory, $"{newName}-{counter}{extension}");
            counter++;
          }

          File.Move(file, newPath);
        }
        catch (Exception ex)
        {
          Console.WriteLine($"Rename failed: {file} - {ex.Message}");
        }
      }

    }

    private void RenameAddText(string folder)
    {
      string textToAdd = txtAddText.Text;

      if (string.IsNullOrWhiteSpace(textToAdd))
        return;

      var files = Directory.GetFiles(folder, "*", System.IO.SearchOption.AllDirectories);

      foreach (var file in files)
      {
        try
        {
          string directory = Path.GetDirectoryName(file);
          string nameOnly = Path.GetFileNameWithoutExtension(file);
          string extension = Path.GetExtension(file).ToLower();

          // ⛔ Skip unwanted files
          if (extension == ".docx" || extension == ".doc" || extension == ".zip")
            continue;

          // Skip if already has text
          if (nameOnly.StartsWith(textToAdd + "-"))
            continue;

          string newName = textToAdd + "-" + nameOnly;
          newName = newName.Replace("--", "-");

          string newPath = Path.Combine(directory, newName + extension);
          if (file == newPath)
            continue;
          File.Move(file, newPath);
        }
        catch { }
      }


    }

    private void RenameNumbers(string folder)
    {
      // Check if the directory exists
      if (Directory.Exists(folder))
      {
        // Get all files in the directory
        string[] files = System.IO.Directory.GetFiles(folder, "*", System.IO.SearchOption.AllDirectories);
        int i = 0;
        // Iterate through each file
        foreach (string filePath in files)
        {

          try
          {
            // Extract the file name and fileExtension           
            string fileExtension = System.IO.Path.GetExtension(filePath);
            string newFileName = System.IO.Path.GetFileName(filePath);
            fileExtension = fileExtension.ToLower();

            if (fileExtension != ".pdf" &&
            fileExtension != ".docx" &&
            fileExtension != ".doc" &&
            fileExtension != ".zip")
            {
              // Construct the new file name
              newFileName = string.Concat(newFileName, "-", ++i);
              newFileName = newFileName.Replace(fileExtension, "");
              newFileName += fileExtension;
              newFileName = newFileName.ToLower();
              newFileName = newFileName.Replace("--", "-");

              //capitalize any main or lead items
              newFileName = renameMAINLEAD(newFileName);

              // Combine the new file name with the original directory
              string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);

              // Rename the file
              if (!File.Exists(newFilePath))
              {
                File.Move(filePath, newFilePath);
              }
            }
          }
          catch (Exception ex)
          {
            MessageBox.Show($"Error renaming file: {ex.Message}");
          }
        }
      }
      else
      {
        MessageBox.Show("Downloads directory does not exist.");
      }
    }

    //Rename the files but remove all the text first,
    //rename it with numbers only. need fix for subfolders
    private void RenameNumbersOnly(String folder)
    {
      if (!Directory.Exists(folder))
      {
        MessageBox.Show("Directory does not exist.");
        return;
      }

      var files = Directory.GetFiles(folder)
                           .OrderBy(f => f)
                           .ToList();

      int i = 0;

      foreach (string filePath in files)
      {
        try
        {
          string fileExtension = Path.GetExtension(filePath).ToLower();

          // Skip unwanted extensions if needed
          if (fileExtension == ".docx" || fileExtension == ".doc" || fileExtension == ".zip")
            continue;

          string newFileName = $"-{++i}{fileExtension}";
          newFileName = newFileName.Replace("--", "-");
          string newFilePath = Path.Combine(folder, newFileName);

          if (!File.Exists(newFilePath))
          {
            File.Move(filePath, newFilePath);
          }
        }
        catch (Exception ex)
        {
          MessageBox.Show($"Error renaming file: {ex.Message}");
        }
      }
    }

    private void renameSpecial(string folder)
    {

      // Check if the directory exists
      if (System.IO.Directory.Exists(folder))
      {
        // Get all files in the directory
        string[] files = System.IO.Directory.GetFiles(folder, "*", System.IO.SearchOption.AllDirectories);

        // Iterate through each file
        foreach (string filePath in files)
        {
          try
          {
            // Extract the file name and fileExtension
            string fileName = System.IO.Path.GetFileName(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            fileExtension = fileExtension.ToLower();
            if (fileExtension != ".docx" && fileExtension != ".doc" && fileExtension != ".zip" && (fileExtension == ".jpg" ||
            fileExtension == ".jpeg" || fileExtension == "heic" || fileExtension == ".png" || fileExtension == ".pdf"))
            {
              string newFileName = fileName.Replace(fileExtension, "");

              newFileName = newFileName.Replace("&", "_");
              newFileName = newFileName.Replace("img", "");
              newFileName = newFileName.Replace("images", "");
              newFileName = newFileName.Replace("image", "");
              newFileName = newFileName.Replace("good", "");
              newFileName = newFileName.Replace("new site", "");
              newFileName = newFileName.Replace("for website", "");
              newFileName = newFileName.Replace("website", "");
              newFileName = newFileName.Replace("web", "");
              newFileName = newFileName.Replace("dsc", "");
              newFileName = newFileName.Replace("unknown", "");
              newFileName = newFileName.Replace("unnamed", "");
              newFileName = newFileName.Replace("untitled", "");
              newFileName = newFileName.Replace("screenshot", "");
              newFileName = newFileName.Replace("picture", "");
              newFileName = newFileName.Replace("photo", "");
              newFileName = newFileName.Replace("thumbnail", "");
              newFileName = newFileName.Replace(" ", "_");
              newFileName = newFileName.Replace(".", "");
              newFileName = newFileName.Replace("(", "");
              newFileName = newFileName.Replace(";", "");
              newFileName = newFileName.Replace("docx", "");
              newFileName = newFileName.Replace("pptx", "");
              newFileName = newFileName.Replace("jpeg", "");
              newFileName = newFileName.Replace("png", "");
              newFileName = newFileName.Replace("copy of", "");
              newFileName = newFileName.Replace("copy", "");
              newFileName = newFileName.Replace("google_docs", "");
              newFileName = newFileName.Replace("processed", "");
              newFileName = newFileName.Replace(")", "");
              newFileName = newFileName.Replace("[", "");
              newFileName = newFileName.Replace("]", "");
              newFileName = newFileName.Replace("’", "");
              newFileName = newFileName.Replace("_-_", "-");
              newFileName = newFileName.Replace("~", "");
              newFileName = newFileName.Replace(",", "");
              newFileName = newFileName.Replace("‘", "");
              newFileName = newFileName.Replace("#", "");
              newFileName = newFileName.Replace("'", "");
              newFileName = newFileName.Replace("-_", "-");
              newFileName = newFileName.Replace("_-", "-");
              newFileName = newFileName.Replace("--", "-");
              newFileName = newFileName.Replace("__", "_");
              newFileName = newFileName.Replace("__", "_");
              newFileName = newFileName.Replace("--", "-");

              newFileName = newFileName.ToLower();

              fileExtension = fileExtension.ToLower();
              fileExtension = fileExtension.Replace("jpeg", "jpg");

              // ---------------------------------------***capitalize any main or lead items****

              newFileName = renameMAINLEAD(newFileName);

              newFileName = newFileName + fileExtension;

              if (chkMaxLength45.IsChecked == true)
              {
                if (newFileName.Length > 45 && !newFileName.Contains(".pdf"))
                {
                  newFileName = newFileName.Replace("_", "");
                  newFileName = newFileName.Replace("-", "");

                  if (newFileName.Length > 45)
                  {
                    MessageBox.Show($"Warning: Filename too long - newfilename {newFileName} is {newFileName.Length} chars. Must be under 45. ");
                  }
                }
              }

              // Combine the new file name with the original directory
              string newFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath), newFileName);

              // Rename the file
              System.IO.File.Move(filePath, newFilePath);
            }
          }
          catch
          {
          }
        }
      }
      else
      {
        MessageBox.Show("Downloads directory does not exist.");
      }
    }

    private void btnTitleCase_Click(object sender, RoutedEventArgs e)
    {
      renameVBTitleCase(downloadsPath);
    }
    private void renameVBTitleCase(string folder)
    {
      this.Title = "Micro Renamer for Windows - Syntax Communications";
      string folderPath = folder;

      if (System.IO.Directory.Exists(folderPath))
      {
        string[] files = System.IO.Directory.GetFiles(folderPath, "*", System.IO.SearchOption.AllDirectories);

        foreach (string filePath in files)
        {
          try
          {

            string fileName = System.IO.Path.GetFileName(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            fileExtension = fileExtension.ToLower();

            if (fileExtension == ".pdf")
            {
              // Construct the new file name
              string newFileName = fileName.Replace(fileExtension, "");
              newFileName = newFileName.Replace("&", "_");
              newFileName = newFileName.Replace("’", "_");
              newFileName = newFileName.Replace(".", "_");
              newFileName = newFileName.Replace("docx", "");
              newFileName = newFileName.Replace("pptx", "");
              newFileName = newFileName.Replace("copy", "");
              newFileName = newFileName.Replace("(", "_");
              newFileName = newFileName.Replace(")", "_");
              newFileName = newFileName.Replace("]", "_");
              newFileName = newFileName.Replace("‘", "_");
              newFileName = newFileName.Replace(";", "_");
              newFileName = newFileName.Replace("[", "_");
              newFileName = newFileName.Replace(",", "_");
              newFileName = newFileName.Replace("#", "_");
              newFileName = newFileName.Replace("~", "_");
              newFileName = newFileName.Replace("'", "_");
              newFileName = newFileName.Replace("__", "_");
              newFileName = newFileName.Replace("__", "_");
              newFileName = newFileName.Replace("_", " ");
              newFileName = newFileName.Replace("  ", " ");
              newFileName = newFileName.Replace("  ", " ");
              newFileName = newFileName.ToLower();
              TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
              newFileName = textInfo.ToTitleCase(newFileName);
              newFileName = newFileName.Replace("Etpto", "ETPTO");
              newFileName = newFileName.Replace("Septo", "SEPTO");
              newFileName = newFileName.Replace("Sepfc", "SEPFC");
              newFileName = newFileName.Replace("Esm", "ESM");
              newFileName = newFileName.Replace("Cyo", "CYO");
              newFileName = newFileName.Replace("Sscm", "SSCM");
              newFileName = newFileName.Replace("ESMsa", "ESMSA");
              newFileName = newFileName.Replace("Smpta", "SMPTA");
              newFileName = newFileName.Replace("Esmsa", "ESMSA");
              newFileName = newFileName.Replace("Neysa", "NEYSA");
              newFileName = newFileName.Replace("Scope", "SCOPE");
              newFileName = newFileName.Replace("Stem", "STEM");
              newFileName = newFileName.Replace("Dp", "DP");
              newFileName = newFileName.Replace("DPac", "DPAC");
              newFileName = newFileName.Replace("1St", "1st");
              newFileName = newFileName.Replace("2Nd", "2nd");
              newFileName = newFileName.Replace("3Rd", "3rd");
              newFileName = newFileName.Replace("4Th", "4th");
              newFileName = newFileName.Replace("5Th", "5th");
              newFileName = newFileName.Replace("6Th", "6th");
              newFileName = newFileName.Replace("7Th", "7th");
              newFileName = newFileName.Replace("8Th", "8th");
              newFileName = newFileName.Replace("9Th", "9th");
              newFileName = newFileName.Replace("10Th", "10th");
              newFileName = newFileName.Replace("11Th", "11th");
              newFileName = newFileName.Replace("12Th", "12th");
              newFileName = newFileName.Replace("Jv", "JV");
              newFileName = newFileName.Replace("Parp", "PARP");

              newFileName = newFileName.Replace(GetTodaysDate(), ""); //remove the date if added by accident
              //newFileName = newFileName.Replace(" Pdf", "");
              fileExtension = fileExtension.ToLower();

              newFileName = newFileName + fileExtension;
              // Combine the new file name with the original directory
              string newFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath), newFileName);

              // Rename the file
              System.IO.File.Move(filePath, newFilePath);
            }
          }
          catch
          {

          }
        }
      }
      else
      {
        MessageBox.Show("Downloads directory does not exist.");
      }
    }

    private string renameMAINLEAD(string str)
    {
      str = str.Replace("main", "MAIN");
      str = str.Replace("lead", "LEAD");
      str = str.Replace("slider", "SLIDER");
      str = str.Replace("slideshow", "SLIDESHOW");
      str = str.Replace("cover", "COVER");
      str = str.Replace("gallery", "GALLERY");
      str = str.Replace("news", "NEWS");
      str = str.Replace("feature", "FEATURE");
      return str;
    }   
    private void RenameShortFiles(string folderPath)
    {
      // Get Word doc filename
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
      string wordFile = Directory.GetFiles(
      folderPath,
      "*.docx",
      System.IO.SearchOption.AllDirectories
      ).FirstOrDefault();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
      string wordFileName = null;

      if (!string.IsNullOrEmpty(wordFile))
      {
        wordFileName = Path.GetFileNameWithoutExtension(wordFile);

        foreach (char c in Path.GetInvalidFileNameChars())
          wordFileName = wordFileName.Replace(c.ToString(), "");
      }

      var files = Directory.GetFiles(folderPath, "*.*", System.IO.SearchOption.AllDirectories);

      foreach (var file in files)
      {
        try
        {
          string directory = Path.GetDirectoryName(file);
          string fileName = Path.GetFileNameWithoutExtension(file);
          string extension = Path.GetExtension(file);

          string datePart = "";
          string remainder = fileName;

          // If it starts with a date like 040626-
          if (System.Text.RegularExpressions.Regex.IsMatch(fileName, @"^\d{6}-"))
          {
            datePart = fileName.Substring(0, 7);
            remainder = fileName.Substring(7);
          }

          string numberSuffix = "";

          // ✅ FIXED LOGIC
          if (string.IsNullOrWhiteSpace(remainder))
          {
            numberSuffix = "0"; // default if nothing after date
          }
          else if (System.Text.RegularExpressions.Regex.IsMatch(remainder, @"^\d+$"))
          {
            numberSuffix = remainder; // any length number
          }
          else
          {
            continue; // skip anything with letters
          }

          string baseName = !string.IsNullOrWhiteSpace(wordFileName)
              ? wordFileName
              : "story";

          string newName = datePart + baseName + "-" + numberSuffix;

          string newPath = Path.Combine(directory, newName + extension);

          int counter = 1;
          while (File.Exists(newPath))
          {
            newPath = Path.Combine(directory, $"{datePart}{baseName}-{numberSuffix}-{counter}{extension}");
            counter++;
          }

          if (file != newPath)
          {
            File.Move(file, newPath);
          }
        }
        catch (Exception ex)
        {
          Console.WriteLine($"Rename failed: {file} - {ex.Message}");
        }
      }


    }

    private void btnRecycle_Click(object sender, RoutedEventArgs e)
    {
      DeleteFilesInDirectory(downloadsPath);
      DeleteEmptyFoldersInDownloads();
    }
    private void DeleteFilesInDirectory(string directoryPath)
    {
      this.Title = "Micro Renamer for Windows - Syntax Communications";
      try
      {
        // Check if the directory exists
        if (System.IO.Directory.Exists(directoryPath))
        {
          // Delete all files in the directory
          foreach (string filePath in System.IO.Directory.GetFiles(directoryPath))
          {
            //File.Delete(filePath);
            FileSystem.DeleteFile(filePath,
                UIOption.OnlyErrorDialogs,
                RecycleOption.SendToRecycleBin);
          }
          // Recursively delete files in subdirectories
          foreach (string subDirectoryPath in System.IO.Directory.GetDirectories(directoryPath))
          {
            DeleteFilesInDirectory(subDirectoryPath);
          }
        }
      }
      catch
      {
        MessageBox.Show("Please close any open files to delete them");
        DeleteFilesInDirectory(directoryPath);
      }
    }
    private void DeleteEmptyFoldersInDownloads()
    {



      DeleteEmptyDirectories(downloadsPath);


    }
    private void DeleteEmptyDirectories(string path)
    {
      foreach (var directory in Directory.GetDirectories(path))
      {
        // First clean subdirectories
        DeleteEmptyDirectories(directory);

        // Then check if current directory is empty
        if (!Directory.EnumerateFileSystemEntries(directory).Any())
        {
          try
          {
            Directory.Delete(directory);
          }
          catch (Exception ex)
          {
            Console.WriteLine($"Could not delete {directory}: {ex.Message}");
          }
        }
      }
    }//end last function
    private void DeleteMacOSXFolders(string rootPath)
    {
      try
      {
        var dirs = Directory.GetDirectories(
            rootPath,
            "__MACOSX",
            System.IO.SearchOption.AllDirectories
        );

        foreach (var dir in dirs)
        {
          try
          {
            FileSystem.DeleteDirectory(
                dir,
                UIOption.OnlyErrorDialogs,
                RecycleOption.SendToRecycleBin
            );
          }
          catch (Exception ex)
          {
            Console.WriteLine($"Failed to delete {dir}: {ex.Message}");
          }
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine($"Error scanning for __MACOSX folders: {ex.Message}");
      }
    }

    private void btnCopyNotepad1_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad1.SelectAll();
      txtNotepad1.Copy();
    }
    private void btnCopyNotepad2_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad2.SelectAll();
      txtNotepad2.Copy();
    }
    private void btnCopyNotepad3_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad3.SelectAll();
      txtNotepad3.Copy();
    }
    private void btnCopyNotepad4_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad4.SelectAll();
      txtNotepad4.Copy();
    }


    private void btnCopyPresetClipboard0_Click(object sender, RoutedEventArgs e)
    {
      txtPresetClipboard0.SelectAll();
      txtPresetClipboard0.Copy();
    }
    private void btnCopyPresetClipboard1_Click(object sender, RoutedEventArgs e)
    {
      //copy the text in the box 1 to the clipboard
      txtPresetClipboard1.SelectAll();
      txtPresetClipboard1.Copy();
    }
    private void btnCopyPresetClipboard2_Click(object sender, RoutedEventArgs e)
    {
      //copy the text in the box 2 to the clipboard
      txtPresetClipboard2.SelectAll();
      txtPresetClipboard2.Copy();
    }
    private void btnCopyPresetClipboard3_Click(object sender, RoutedEventArgs e)
    {
      //copy the text in the box 1 to the clipboard
      txtPresetClipboard3.SelectAll();
      txtPresetClipboard3.Copy();
    }
    private void btnCopyPresetClipboard4_Click(object sender, RoutedEventArgs e)
    {
      //copy the text in the box 1 to the clipboard
      txtPresetClipboard4.SelectAll();
      txtPresetClipboard4.Copy();
    }

    private void btnPasteNotepad1_Click(object sender, RoutedEventArgs e)
    {
      string cb = Clipboard.GetText();
      txtNotepad1.Text = cb;
    }
    private void btnPasteNotepad2_Click(object sender, RoutedEventArgs e)
    {
      string cb = Clipboard.GetText();
      txtNotepad2.Text = cb;
    }
    private void btnPasteNotepad3_Click(object sender, RoutedEventArgs e)
    {

      string cb = Clipboard.GetText();
      txtNotepad3.Text = cb;
    }

    private void btnClearNotepad1_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad1.Clear();
    }
    private void btnClearNotepad2_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad2.Clear();
    }
    private void btnClearNotepad3_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad3.Clear();
    }
    private void btnClearNotepad4_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad4.Clear();
    }

    //btn For the WORD GET TITLE TEXT
    private async void GetTitleTextWord_Click(object sender, RoutedEventArgs e)
    {

      ExtractAllZipsInDownloads();

      string filePath = null;

      try
      {
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
        filePath = Directory.GetFiles(
            downloadsPath,
            "*.docx",
            System.IO.SearchOption.AllDirectories
        ).FirstOrDefault();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
      }
      catch (UnauthorizedAccessException)
      {
        MessageBox.Show("Access denied to some folders.");
      }

      if (filePath == null)
      {
        txtNotepad1.Text = "";
        txtNotepad2.Text = "";
      }
      else
      {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
        {
          var paragraphs = doc.MainDocumentPart.Document.Body
              .Elements<Paragraph>()
              .Select(p => p.InnerText.Trim())
              .Where(p => !string.IsNullOrWhiteSpace(p))
              .ToList();

          if (paragraphs.Count > 0)
          {
            int startIndex = 0;

            int syntaxIndex = paragraphs.FindIndex(p => p.ToLower().Contains("syntaxny.com"));
            if (syntaxIndex != -1)
              startIndex = syntaxIndex + 1;

            var filtered = paragraphs.Skip(startIndex).ToList();

            int titleIndex = filtered.FindIndex(p =>
                p.Length > 20 &&
                !p.Contains("@") &&
                !p.Contains("www.") &&
                !p.ToUpper().Equals(p) &&
                !System.Text.RegularExpressions.Regex.IsMatch(p, @"^\d")
            );

            if (titleIndex == -1)
              titleIndex = 0;

            txtNotepad1.Text = filtered[titleIndex];
            txtNotepad2.Text = string.Join(Environment.NewLine, filtered.Skip(titleIndex + 1));
          }
        }
      }

      //sometimes files contain the press release date + other info
      if (txtNotepad2.Text.Contains("syntaxny"))
      {
        SplitTextboxContent();
      }
      //if AE sends MacOS files. 
      DeleteMacOSXFolders(downloadsPath);

      //shrink images not .heic to 800x600
      ProcessImagesInDownloads();

      RenameFiles();
      RenameFiles();
      await GenerateAltText();
    }
    private void SplitTextboxContent()
    {
      var lines = txtNotepad2.Text
      .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
      .Select(l => l.Trim())
      .Where(l => !string.IsNullOrWhiteSpace(l))
      .ToList();


      if (lines.Count == 0)
        return;

      int titleIndex = lines.FindIndex(l =>
          l.Length > 20 &&                                // real sentence
          !l.Contains("@") &&                             // not email
          !l.Contains("www.") &&                          // not website
          !l.ToUpper().Equals(l) &&                       // not ALL CAPS
          !System.Text.RegularExpressions.Regex.IsMatch(l, @"^\d+$") // not numbers
      );

      if (titleIndex == -1)
        titleIndex = 0;

      txtNotepad1.Text = lines[titleIndex];

      txtNotepad2.Text = string.Join(Environment.NewLine,
          lines.Skip(titleIndex + 1));


    }
    private void ProcessImagesInDownloads()
    {
      this.Title = "Micro Renamer for Windows - Syntax Communications";

      this.Title += " - Processing...";



      string jpegPath = Path.Combine(downloadsPath, "JPEG");
      if (Directory.Exists(jpegPath))
      {
        try
        {
          FileSystem.DeleteDirectory(
              jpegPath,
              UIOption.OnlyErrorDialogs,
              RecycleOption.SendToRecycleBin
          );
        }
        catch (Exception ex)
        {
          MessageBox.Show($"Could not delete folder: {ex.Message}");
        }
      }


      // 🚫 Check for HEIC files first
      var heicFiles = Directory.GetFiles(downloadsPath, "*.heic", System.IO.SearchOption.AllDirectories);
      if (heicFiles.Length > 0)
      {
        MessageBox.Show("HEIC files detected. Please use photoshop to process files.");
        return;
      }

      var imageExtensions = new[] { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.gif" };

      var files = imageExtensions
          .SelectMany(ext => Directory.GetFiles(
              downloadsPath,
              ext,
              System.IO.SearchOption.AllDirectories))
          .ToList();

      foreach (var inputPath in files)
      {
        try
        {
          if (inputPath.Contains("\\JPEG\\")) continue;

          using (var image = ISImage.Load(inputPath))
          {
            image.Mutate(x => x.Resize(new ResizeOptions
            {
              Mode = ISResizeMode.Max,
              Size = new ISSize(800, 600)
            }));

            string relativePath = Path.GetRelativePath(downloadsPath, inputPath);
            string outputPath = Path.Combine(jpegPath, Path.ChangeExtension(relativePath, ".jpg"));
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            image.Save(outputPath, new JpegEncoder
            {
              Quality = 95
            });
          }
        }
        catch (Exception ex)
        {
          Console.WriteLine($"Skipped {inputPath}: {ex.Message}");
        }
      }

      this.Title = "Micro Renamer for Windows - Syntax Communications";

      this.Title += " - Processing Complete...";

    } //end last function
    private async Task GenerateAltText()
    {
      string apiKeyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "apikey.txt");

      string apiKey = File.Exists(apiKeyPath)
          ? File.ReadAllText(apiKeyPath).Trim()
          : Environment.GetEnvironmentVariable("OPENAI_API_KEY");

      if (string.IsNullOrWhiteSpace(apiKey))
      {
        MessageBox.Show("API key not found. Dont forget to add a file named 'apikey.txt' in your project folder with openai key");
        return;
      }


      string jpegPath = Path.Combine(downloadsPath, "JPEG");

      if (!Directory.Exists(jpegPath))
      {
        MessageBox.Show("JPEG folder not found.");
        return;
      }

      var files = Directory.GetFiles(jpegPath, "*.jpg", System.IO.SearchOption.AllDirectories);

      List<string> lines = new List<string>();

      using (var http = new HttpClient())
      {
        http.DefaultRequestHeaders.Authorization =
            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

        foreach (var file in files)
        {
          try
          {
            byte[] imageBytes = File.ReadAllBytes(file);
            string base64 = Convert.ToBase64String(imageBytes);

            var requestBody = new
            {
              model = "gpt-4o-mini",
              messages = new object[]
                {
                    new {
                        role = "user",
                        content = new object[]
                        {
                            new { type = "text", text = "Generate a short SEO-friendly alt text for this image." },
                            new {
                                type = "image_url",
                                image_url = new {
                                    url = $"data:image/jpeg;base64,{base64}"
                                }
                            }
                        }
                    }
                }
            };

            string json = System.Text.Json.JsonSerializer.Serialize(requestBody);

            var response = await http.PostAsync(
                "https://api.openai.com/v1/chat/completions",
                new StringContent(json, System.Text.Encoding.UTF8, "application/json")
            );

            string result = await response.Content.ReadAsStringAsync();

            using var doc = System.Text.Json.JsonDocument.Parse(result);
            string altText = doc.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString();

            altText = altText.Trim().Trim('"');

            string fileName = Path.GetFileName(file);

            lines.Add(fileName);
            lines.Add(altText);
            lines.Add("");
          }
          catch (Exception ex)
          {
            lines.Add($"{Path.GetFileName(file)}: ERROR");
          }
        }
      }

      txtNotepad4.Text = string.Join("\r\n", lines);

    }
    private void ExtractAllZipsInDownloads()
    {


      var zipFiles = Directory.GetFiles(downloadsPath, "*.zip");

      foreach (var zip in zipFiles)
      {
        try
        {
          // Extract to same folder
          ZipFile.ExtractToDirectory(zip, downloadsPath, true);
        }
        catch (Exception ex)
        {
          Console.WriteLine($"Failed to extract {zip}: {ex.Message}");
        }
      }
    } //end function

    private void btnCAPS_Click(object sender, RoutedEventArgs e)
    {
      if (textMode == 0)
      {
        // CAPS
        txtNotepad4.Text = txtNotepad4.Text.ToUpper();
        btnCAPS.Content = "lower";
        textMode = 1;
      }
      else if (textMode == 1)
      {
        // lower
        txtNotepad4.Text = txtNotepad4.Text.ToLower();
        btnCAPS.Content = "Title";
        textMode = 2;
      }
      else
      {
        // Title Case
        var textInfo = CultureInfo.CurrentCulture.TextInfo;
        txtNotepad4.Text = textInfo.ToTitleCase(txtNotepad4.Text.ToLower());

        btnCAPS.Content = "CAPS";
        textMode = 0;
      }
    }
    private void btnGetBullet_Click(object sender, RoutedEventArgs e)
    {
      string link = txtHTMLLink.Text.Trim();
      string text = txtHTMLText.Text.Trim();

      if (string.IsNullOrWhiteSpace(link) || string.IsNullOrWhiteSpace(text))
        return;

      string html = $"<li><a href=\"{link}\">{text}</a></li>";

      txtNotepad4.Text = html;
    }
    private void btnBullets_DL_Click(object sender, RoutedEventArgs e)
    {
      LoadHtmlListToTextBox(downloadsPath);
    }
    private void btnTitlesDL_Click(object sender, RoutedEventArgs e)
    {
      LoadFileNamesToTextBox(downloadsPath);
    }
    private void LoadFileNamesToTextBox(string folder)
    {


      if (!Directory.Exists(folder))
      {
        MessageBox.Show("Directory does not exist.");
        return;
      }

      var textInfo = CultureInfo.CurrentCulture.TextInfo;

      var fileNames = Directory.GetFiles(folder)
                               .OrderBy(f => f)
                               .Select(file =>
                               {
                                 string name = Path.GetFileNameWithoutExtension(file);
                                 name = name.Replace(GetTodaysDate(), "");
                                 // Clean + Title Case
                                 name = name.Replace("_", " ")
                                          .Replace("-", " ");

                                 name = textInfo.ToTitleCase(name.ToLower());


                                 return name;
                               });

      txtNotepad4.Text = string.Join(Environment.NewLine, fileNames);
    }
    private void LoadHtmlListToTextBox(string folder)
    {
      if (!Directory.Exists(folder))
      {
        MessageBox.Show("Directory does not exist.");
        return;
      }

      var textInfo = CultureInfo.CurrentCulture.TextInfo;

      var listItems = Directory.GetFiles(folder)
                               .OrderBy(f => f)
                               .Select(file =>
                               {
                                 string name = Path.GetFileNameWithoutExtension(file);

                                 // Remove today's date
                                 string today = GetTodaysDate();
                                 if (!string.IsNullOrEmpty(today))
                                   name = name.Replace(today, "");

                                 // Clean separators
                                 name = name.Replace("_", " ")
                                          .Replace("-", " ");

                                 // Remove extra spaces
                                 name = string.Join(" ", name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

                                 // Title Case
                                 name = textInfo.ToTitleCase(name.ToLower());

                                 // Convert to HTML-safe text
                                 name = System.Net.WebUtility.HtmlEncode(name);

                                 return $"  <li>{name}</li>";
                               });

      string html =
          "<ul>\r\n" +
          string.Join("\r\n", listItems) +
          "\r\n</ul>";

      txtNotepad4.Text = html;
    }

  }//end of form

}//end namespace