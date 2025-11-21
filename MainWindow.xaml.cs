using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using Microsoft.VisualBasic.FileIO;

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

    private void getFolderPath(TextBox txtBox) {

      //This function sets the selected folder path as the text in the directory text box

      Microsoft.Win32.OpenFolderDialog dialog = new();

      dialog.Multiselect = false;
      dialog.Title = "Select a folder";

      // Show open folder dialog box
      bool? result = dialog.ShowDialog();

      string fullPathToFolder = "";
      // Process open folder dialog box results
      if (result == true)
      {
        // Get the selected folder
        fullPathToFolder = dialog.FolderName;
      }

      txtBox.Text = fullPathToFolder;
    }

    private void RenameFiles(TextBox txtFolder){

      if (chkAddDate.IsChecked == true || chkAddText.IsChecked == true)
      {
        renameWithDateAndText(txtFolder.Text);
      }
      if (chkRemoveSpecial.IsChecked == true)
      {
        renameSpecial(txtFolder.Text);
      }
      if (chkNumberItems.IsChecked == true)
      {
        renameNumbers(txtFolder.Text);
      }
    }

    private void renameWithDateAndText(string folder){
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
            // Extract the file name and extension           
            string fileExtension = System.IO.Path.GetExtension(filePath);
            fileExtension = fileExtension.ToLower();
            if (fileExtension != ".docx" && fileExtension != ".doc" && fileExtension != ".zip")
            {
              string a = "", b = "";
              string todaysDate = "";

              //Add text if the option is selected
              if (chkAddText.IsChecked == true)
              {
                if (txtAddText.Text != "")
                {
                  a = txtAddText.Text;
                  b = "-";
                }
              }

              //Add date if the option is selected
              if (chkAddDate.IsChecked == true)
              {

                todaysDate = string.Concat(DateTime.Now.Month.ToString("D2"), DateTime.Now.ToString("dd"), DateTime.Now.ToString("yy"), "-");
              }
              string newFileName = string.Concat(todaysDate, a, b, System.IO.Path.GetFileName(filePath));

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
      txtAddText.Clear();
    }
    // ---------------------------------------***capitalize any main or lead items****

    private string renameMAINLEAD(string str) {
      str = str.Replace("main", "MAIN");
      str = str.Replace("lead", "LEAD");
      str = str.Replace("slider", "SLIDER");
      str = str.Replace("slideshow", "SLIDESHOW");
      str = str.Replace("cover", "COVER");
      str = str.Replace("gallery", "GALLERY");
      str = str.Replace("news", "NEWS");
      return str;
    }
    private void renameSpecial(string folder) {     

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
            // Extract the file name and extension
            string fileName = System.IO.Path.GetFileName(filePath);
            string fileExtension = System.IO.Path.GetExtension(filePath);
            fileExtension = fileExtension.ToLower();
            if (fileExtension != ".docx" && fileExtension != ".doc" && fileExtension != ".zip" && (fileExtension == ".jpg" ||
            fileExtension == ".jpeg" || fileExtension == "heic" || fileExtension == ".png" || fileExtension == ".pdf"))
            {
              string newFileName = fileName.Replace(fileExtension, "");

              newFileName = newFileName.Replace("&", "_");
              newFileName = newFileName.Replace("img", "");
              newFileName = newFileName.Replace("image", "");
              newFileName = newFileName.Replace("dsc", "");
              newFileName = newFileName.Replace("unknown", "");
              newFileName = newFileName.Replace("unnamed", "");
              newFileName = newFileName.Replace("untitled", "");
              newFileName = newFileName.Replace("screenshot", "");
              newFileName = newFileName.Replace("picture", "");
              newFileName = newFileName.Replace("photo", "");
              newFileName = newFileName.Replace(" ", "_");
              newFileName = newFileName.Replace(".", "");
              newFileName = newFileName.Replace("(", "");
              newFileName = newFileName.Replace(";", "");
              newFileName = newFileName.Replace("docx", "");
              newFileName = newFileName.Replace("pptx", "");
              newFileName = newFileName.Replace("jpeg", "");
              newFileName = newFileName.Replace("png", "");
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
            // Extract the file name and extension           
            string fileExtension = System.IO.Path.GetExtension(filePath);
            string newFileName = System.IO.Path.GetFileName(filePath);
            fileExtension = fileExtension.ToLower();

            if (fileExtension != ".pdf" && fileExtension != ".docx" && fileExtension != ".doc" && fileExtension != ".zip")
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
              newFileName = newFileName.Replace("Jv", "JV");
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


    private void btnOpenFolderDir1_Click(object sender, RoutedEventArgs e)
    {
      //gets the selected folder path
      getFolderPath(txtDirectory1);
    }

    private void btnOpenFolderDir2_Click(object sender, RoutedEventArgs e)
    {
      //gets the selected folder path
      getFolderPath(txtDirectory2);
    }

    private void btnRenameDir1_Click(object sender, RoutedEventArgs e)
    {
      RenameFiles(txtDirectory1);
    }

    private void btnRenameDir2_Click(object sender, RoutedEventArgs e)
    {
      RenameFiles(txtDirectory2);
    }

    private void btnTitleCaseDir1_Click(object sender, RoutedEventArgs e)
    {
      renameVB(txtDirectory1.Text);
    }

    private void btnTitleCaseDir2_Click(object sender, RoutedEventArgs e)
    {
      renameVB(txtDirectory2.Text);
    }

    private void btnRecycleDir1_Click(object sender, RoutedEventArgs e)
    {
      DeleteFilesInDirectory(txtDirectory1.Text);
    }
    private void btnRecycleDir2_Click(object sender, RoutedEventArgs e)
    {
      DeleteFilesInDirectory(txtDirectory2.Text);
    }

    private void btnRecycleBoth_Click(object sender, RoutedEventArgs e)
    {
      DeleteFilesInDirectory(txtDirectory1.Text);
      DeleteFilesInDirectory(txtDirectory2.Text);
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
  }
}