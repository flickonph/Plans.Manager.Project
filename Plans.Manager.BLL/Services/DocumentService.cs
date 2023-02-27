namespace Plans.Manager.BLL.Services;

public static class DocumentService
{
    public static bool CreateAndSavePlan(List<string> plxFiles, int course, int semester, int weeks, int year, string dirToSave)
    {
        var excelHandler = new ExcelHandler(plxFiles, year);
        var plansPackage = excelHandler.PlansPackage(course, semester, weeks, year);
        var excelFileInfo = new FileInfo($@"{dirToSave}\РП {course}к.({(course * 2 - 2) + semester} сем.) {year}.xlsx");
        if (plansPackage.Workbook.Worksheets.Count == 0)
            return true;

        switch (excelFileInfo.Exists)
        {
            case true:
                try
                {
                    File.Delete(excelFileInfo.FullName);
                    plansPackage.SaveAs(excelFileInfo);
                    plansPackage.Dispose();
                    return true;
                }
                catch
                {
                    plansPackage.Dispose();
                    return false;
                }

            default:
                plansPackage.SaveAs(excelFileInfo);
                plansPackage.Dispose();
                return true;
        }
    }

    public static bool CreateAndSaveLoad(List<string> plxFiles, int semester, int year, string dirToSave)
    {
        var excelHandler = new ExcelHandler(plxFiles, year);
        var loadPackage = excelHandler.StudyLoad(semester);
        var excelFileInfo = new FileInfo($@"{dirToSave}\Кафедральная нагрузка({semester} сем.) {year}.xlsx");

        switch (excelFileInfo.Exists)
        {
            case true:
                try
                {
                    File.Delete(excelFileInfo.FullName);
                    loadPackage.SaveAs(excelFileInfo);
                    loadPackage.Dispose();
                    return excelFileInfo.Exists;
                }
                catch
                {
                    loadPackage.Dispose();
                    return false;
                }

            default:
                loadPackage.SaveAs(excelFileInfo);
                loadPackage.Dispose();
                break;
        }

        return excelFileInfo.Exists;
    }
    
    public static bool CreateAndSaveDepartments(string path)
    {
        var datHandler = new DatHandler(path);
        const string fileName = @"Config\departments.xml";
        var currentFile = new FileInfo(fileName);
        var xmlDocument = datHandler.DatToXml();

        switch (currentFile.Exists)
        {
            case true:
                try
                {
                    File.Delete(currentFile.FullName);
                    xmlDocument.Save(currentFile.FullName);
                    return true;
                }
                catch
                {
                    return false;
                }

            case false:
                xmlDocument.Save(currentFile.FullName);
                break;
        }

        return currentFile.Exists;
    }
}