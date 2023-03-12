using System.Xml;
using OfficeOpenXml;
using Plans.Manager.BLL.Builders;
using Plans.Manager.BLL.Readers;

namespace Plans.Manager.BLL.Services;

public static class DocumentService
{
    public static bool CreateAndSavePlan(List<string> plxFilesPaths, int selectedCourse, int selectedSemester, int selectedWeeks, int selectedYear, string dirToSave)
    {
        bool status = true;
        PlanBuilder planBuilder = new PlanBuilder(plxFilesPaths, selectedYear);
        ExcelPackage planPackage = planBuilder.PlansExcelPackage(selectedCourse, selectedSemester, selectedWeeks);
        FileInfo planFileInfo = new FileInfo($@"{dirToSave}\РП {selectedCourse}к.({(selectedCourse * 2 - 2) + selectedSemester} сем.) {selectedYear}.xlsx");
        if (planPackage.Workbook.Worksheets.Count == 0)
            return false;

        switch (planFileInfo.Exists)
        {
            case true:
                try
                {
                    File.Delete(planFileInfo.FullName);
                    planPackage.SaveAs(planFileInfo);
                    planPackage.Dispose();
                }
                catch
                {
                    planPackage.Dispose();
                    status = false;
                }
                
                break;

            default:
                planPackage.SaveAs(planFileInfo);
                planPackage.Dispose();
                break;
        }

        return status;
    }

    public static bool CreateAndSaveLoad(List<string> plxFilesPaths, int selectedSemester, int selectedYear, string dirToSave)
    {
        bool status = true;
        LoadBuilder loadBuilder = new LoadBuilder(plxFilesPaths, selectedYear, selectedSemester);
        List<ExcelPackage> loadPackages = loadBuilder.LoadExcelPackages();
        Directory.CreateDirectory($@"{dirToSave}\{selectedYear}");
        foreach (ExcelPackage loadPackage in loadPackages)
        {
            if (loadPackage.Workbook.Worksheets.Count == 0)
                continue;
        
            FileInfo loadFileInfo = new FileInfo($@"{dirToSave}\{selectedYear}\{loadPackage.Workbook.Properties.Comments}.xlsx");
            switch (loadFileInfo.Exists)
            {
                case true:
                    try
                    {
                        File.Delete(loadFileInfo.FullName);
                        loadPackage.SaveAs(loadFileInfo);
                        loadPackage.Dispose();
                    }
                    catch
                    {
                        loadPackage.Dispose();
                        status = false;
                    }

                    break;

                default:
                    loadPackage.SaveAs(loadFileInfo);
                    loadPackage.Dispose();
                    break;
            }
        }

        return status;
    }
    
    public static bool CreateAndSaveDepartments(string path)
    {
        DatReader datReader = new DatReader(path);
        const string fileName = @"Config\Departments.xml";
        FileInfo currentFile = new FileInfo(fileName);
        XmlDocument xmlDocument = datReader.DatToXml();

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