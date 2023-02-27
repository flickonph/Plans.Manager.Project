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
        var dialog = new FolderBrowserDialog();
        dialog.Description = description;
        dialog.UseDescriptionForTitle = true;
        dialog.ShowDialog();
        var path = dialog.SelectedPath;

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
                var dialog = new OpenFileDialog
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
                var dir = SelectFolder("Выберите файл(ы)");
                if (dir == string.Empty)
                {
                    pathToFiles = Array.Empty<string>();
                    break;
                }

                var dirInfo = new DirectoryInfo(dir);
                var fileInfos = dirInfo.GetFiles($"*.{ext}", SearchOption.AllDirectories);
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
        var dialog = new OpenFileDialog
        {
            DefaultExt = $".{ext}",
            Filter = $"{ext.ToUpper()} Files (*.{ext})|*.{ext}",
            Multiselect = true,
            Title = "Desktop"
        };

        dialog.ShowDialog();
        var pathToFile = dialog.FileName;

        return pathToFile;
    }
}