using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic.FileIO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Processing;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using static System.Net.Mime.MediaTypeNames;

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
    private int textMode = 0; // 0 = CAPS, 1 = lower, 2 = Title
    public MainWindow()
    {
      InitializeComponent();
    }
    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      string downloadsPath = Path.Combine(
          Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
          "Downloads"
      );

      txtDirectory1.Text = downloadsPath;
    }

    //The RENAME button
    private void RenameFiles()
    {
      this.Title = "Micro Renamer for Windows - Syntax Communications";

      if (chkRemoveAllText.IsChecked == true)
      {
        renameNumbersOnly(txtDirectory1.Text); //remove all text and number each file (Add text in the "add text box" to add text)
      }
      if (chkAddDate.IsChecked == true || chkAddText.IsChecked == true)
      {
        renameWithDateAndText(txtDirectory1.Text);
      }
      if (chkRemoveSpecial.IsChecked == true)
      {
        renameSpecial(txtDirectory1.Text);
      }
      if (chkNumberItems.IsChecked == true)
      {
        renameNumbers(txtDirectory1.Text);
      }
      RenameShortFiles(txtDirectory1.Text);


    }
    //Rename the files but remove all the text first, rename it with numbers only. not good for subfolders
    private void renameNumbersOnly(String folder)
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
    private void renameWithDateAndText(string folder)
    {
      string todaysDate = GetTodaysDate();
      string textToAdd = txtAddText.Text;

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
            string fileExtension = System.IO.Path.GetExtension(filePath);
            fileExtension = fileExtension.ToLower();
            if (fileExtension != ".docx" && fileExtension != ".doc" && fileExtension != ".zip")
            {


              //Add text if the option is selected
              if (chkAddText.IsChecked == false)
              {
                textToAdd = "";
              }

              //Add date if the option is selected
              if (chkAddDate.IsChecked == false)
              {
                todaysDate = "";
              }
              //Checks if the date is already added and won't add the date twice
              string newFileName = string.Concat(todaysDate, textToAdd, "-", System.IO.Path.GetFileName(filePath).Replace(todaysDate, ""));

              newFileName = newFileName.ToLower();

              // ---------------------------------------***capitalize any main or lead items****
              newFileName = renameMAINLEAD(newFileName);

              if (chkRemoveText.IsChecked == true)
              {
                if (txtRemoveText.Text.Length > 0)
                {
                  string txtremove = txtRemoveText.Text.ToLower();
                  newFileName = newFileName.Replace(txtremove, "");
                }
              }
              // Combine the new file name with the original directory
              string newFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath), newFileName);
newFilePath = newFilePath.Replace("--", "-");
              // Rename the file
              System.IO.File.Move(filePath, newFilePath);
              chkRemoveAllText.IsChecked = false;
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
      txtAddText.Clear();
    }
    // ---------------------------------------***capitalize any main or lead items****
    private string GetTodaysDate()
    {
      return string.Concat(DateTime.Now.Month.ToString("D2"), DateTime.Now.ToString("dd"), DateTime.Now.ToString("yy"), "-");
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

    private void renameNumbers(string folder)
    {
      // Check if the directory exists
      if (System.IO.Directory.Exists(folder))
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
              string newFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath), newFileName);

              // Rename the file
              System.IO.File.Move(filePath, newFilePath);
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

    private void renameVB(string folder)
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

    private void btnRenameDir1_Click(object sender, RoutedEventArgs e)
    {
      RenameFiles();
    }

    private void btnTitleCaseDir1_Click(object sender, RoutedEventArgs e)
    {
      renameVB(txtDirectory1.Text);
    }

    private void btnRecycleBoth_Click(object sender, RoutedEventArgs e)
    {
      DeleteFilesInDirectory(txtDirectory1.Text);
      DeleteEmptyFoldersInDownloads();
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
    private void btnCopyPresetClipboard1_Click(object sender, RoutedEventArgs e)
    {
      //copy the text in the box 1 to the clipboard
      txtPresetClipboard1.SelectAll();
      txtPresetClipboard1.Copy();
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
    private void btnCopyPresetClipboard2_Click(object sender, RoutedEventArgs e)
    {
      //copy the text in the box 2 to the clipboard
      txtPresetClipboard2.SelectAll();
      txtPresetClipboard2.Copy();
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

    private void btnClearNotepad1_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad1.Clear();
    }

    private void btnClearNotepad2_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad2.Clear();
    }

    private void btnCopyPresetClipboard0_Click(object sender, RoutedEventArgs e)
    {
      txtPresetClipboard0.SelectAll();
      txtPresetClipboard0.Copy();
    }

    private void btnPasteNotepad3_Click(object sender, RoutedEventArgs e)
    {

      string cb = Clipboard.GetText();
      txtNotepad3.Text = cb;
    }
    private void btnClearNotepad3_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad3.Clear();
    }

    private void btnUndoDates_Click(object sender, RoutedEventArgs e)
    {
      RemoveDateFromFiles(txtDirectory1.Text, GetTodaysDate());
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

    private void btnTitlesDL_Click(object sender, RoutedEventArgs e)
    {
      LoadFileNamesToTextBox(txtDirectory1.Text);
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

    private void btnBullets_DL_Click(object sender, RoutedEventArgs e)
    {
      LoadHtmlListToTextBox(txtDirectory1.Text);
    }

    private void btnCopyNotepad4_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad4.SelectAll();
      txtNotepad4.Copy();
    }

    private void btnClearNotepad4_Click(object sender, RoutedEventArgs e)
    {
      txtNotepad4.Clear();
    }

    private void btnGetTitleTextWord_Click(object sender, RoutedEventArgs e)
    {

      RenameFiles();

      string downloadsPath = txtDirectory1.Text;

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
        MessageBox.Show("No Word documents found.");
        return;
      }

      using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
      {
        var paragraphs = doc.MainDocumentPart.Document.Body
            .Elements<Paragraph>()
            .Select(p => p.InnerText.Trim())
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToList();

        if (paragraphs.Count == 0)
        {
          MessageBox.Show("Document is empty.");
          return;
        }

        // First paragraph = title
        txtNotepad1.Text = paragraphs[0];

        // Remaining paragraphs = body
        StringBuilder body = new StringBuilder();

        foreach (var p in paragraphs.Skip(1))
        {
          body.AppendLine(p);
        }

        txtNotepad2.Text = body.ToString().Trim();
      }

      if (txtNotepad2.Text.Contains("release")){
        //sometimes files contain the press release date + other info
        SplitTextboxContent(); }

      //if AE sends MacOS files. 
      DeleteMacOSXFolders(downloadsPath);

      //shrink images not .heic to 800x600
      ProcessImagesInDownloads();

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

      // Find first "real" title line
      int titleIndex = lines.FindIndex(l =>
          !l.ToLower().Contains("press release") &&
          !l.ToLower().Contains("for immediate release") &&
          !l.Contains("www.") &&
          l.Length > 20
      );

      if (titleIndex == -1)
        titleIndex = 0;

      // Set title
      txtNotepad1.Text = lines[titleIndex];

      // Set body
      txtNotepad2.Text = string.Join(Environment.NewLine,
          lines.Skip(titleIndex + 1));

      //part2
      lines = txtNotepad2.Text
      .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
      .Select(l => l.Trim())
      .Where(l =>
      !string.IsNullOrWhiteSpace(l) &&
      !l.ToLower().Contains("@syntaxny.com") // remove email line
      )
      .ToList();


      if (lines.Count == 0)
        return;

      // First line = title
      txtNotepad1.Text = lines[0];

      // Remaining = body
      txtNotepad2.Text = string.Join(Environment.NewLine, lines.Skip(1));

    }
        
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

    private void ProcessImagesInDownloads()
    {
      this.Title = "Micro Renamer for Windows - Syntax Communications";

      this.Title += " - Processing...";

      string downloadsPath = Path.Combine(
      Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
      "Downloads"
      );

      string outputRoot = Path.Combine(downloadsPath, "JPEG");


      // 🚫 Check for HEIC files first
      var heicFiles = Directory.GetFiles(downloadsPath, "*.heic", System.IO.SearchOption.AllDirectories);
      if (heicFiles.Length > 0)
      {
        MessageBox.Show("HEIC files detected. Please sue photoshop to process files.");
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
            string outputPath = Path.Combine(outputRoot, Path.ChangeExtension(relativePath, ".jpg"));
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

    private void RenameShortFiles(string folderPath)
    {
//      string downloadsPath = Path.Combine(
//Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
//"Downloads"
//);

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

          // ✅ If it starts with a date, separate it
          if (System.Text.RegularExpressions.Regex.IsMatch(fileName, @"^\d{6}-"))
          {
            datePart = fileName.Substring(0, 7); // "040626-"
            remainder = fileName.Substring(7);
          }

          string numberSuffix = "";

          // If remainder is single digit → move it
          if (remainder.Length == 1 && char.IsDigit(remainder[0]))
          {
            numberSuffix = remainder;
          }
          else if (!string.IsNullOrWhiteSpace(remainder))
          {
            continue; // skip anything else
          }

          string baseName = !string.IsNullOrWhiteSpace(wordFileName)
              ? wordFileName
              : "story";

          string newName = datePart + baseName + "-"+numberSuffix;
         
          string newPath = Path.Combine(directory, newName + extension);

          int counter = 1;
          while (File.Exists(newPath))
          {
            newPath = Path.Combine(directory, $"{datePart}{baseName}{numberSuffix}-{counter}{extension}");
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

    private void DeleteEmptyFoldersInDownloads()
    {
      string downloadsPath = Path.Combine(
      Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
      "Downloads"
      );


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


    //end of form
  }
}
