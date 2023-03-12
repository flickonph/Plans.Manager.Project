using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Plans.Manager.Desktop.Models;

public static class WindowsApi
{
    public static string SelectFolder(string description)
    {
        FolderBrowserDialog dialog = new FolderBrowserDialog();
        dialog.Description = description;
        dialog.UseDescriptionForTitle = true;
        dialog.ShowDialog();
        string path = dialog.SelectedPath;

        return path;
    }

    /// <summary>
    /// Gets path(s) to file(s).
    /// </summary>
    /// <param name="ext"> File(s) extension without dot. </param>
    /// <param name="allInFolder"> Type of selection. </param>
    /// <returns> Files paths. </returns>
    public static IEnumerable<string> SelectFiles(string ext, bool allInFolder)
    {
        IEnumerable<string> pathToFiles;

        switch (allInFolder)
        {
            case false:
                OpenFileDialog dialog = new OpenFileDialog
                {
                    DefaultExt = $".{ext}",
                    Filter = $"{ext.ToUpper()} Files (*.{ext})|*.{ext}",
                    Multiselect = true,
                    Title = "TableManager"
                };

                dialog.ShowDialog();
                pathToFiles = dialog.FileNames;

                break;

            case true:
                string dir = SelectFolder("Выберите файл(ы)");
                if (dir == string.Empty)
                {
                    pathToFiles = Array.Empty<string>();
                    break;
                }

                DirectoryInfo dirInfo = new DirectoryInfo(dir);
                FileInfo[] fileInfos = dirInfo.GetFiles($"*.{ext}", SearchOption.AllDirectories);
                pathToFiles = fileInfos.Select(f => f.FullName);

                break;
        }

        return pathToFiles;
    }

    /// <summary>
    /// Gets path to file.
    /// </summary>
    /// <param name="ext"> File extension without dot. </param>
    /// <returns> Single file path. </returns>
    public static string SelectFile(string ext)
    {
        OpenFileDialog dialog = new OpenFileDialog
        {
            DefaultExt = $".{ext}",
            Filter = $"{ext.ToUpper()} Files (*.{ext})|*.{ext}",
            Multiselect = true,
            Title = "Desktop"
        };

        dialog.ShowDialog();
        string pathToFile = dialog.FileName;

        return pathToFile;
    }
}