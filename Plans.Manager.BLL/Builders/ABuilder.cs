using System.Diagnostics;
using Plans.Manager.BLL.Objects;
using Plans.Manager.BLL.Readers;

namespace Plans.Manager.BLL.Builders;

public abstract class ABuilder
{
    private readonly List<string> _plxFilesPaths;
    private readonly int _selectedYear;

    protected ABuilder(List<string> plxFilesPaths, int selectedYear)
    {
        _plxFilesPaths = plxFilesPaths;
        _selectedYear = selectedYear;
    }
    
    protected static readonly string[] ControlType = { "Зачет", "Зачет с оценкой", "Экзамен", "Курсовая работа", "Курсовой проект" }; // TODO: Move to configuration file instead

    protected List<PlanGroupPair> GetAllPlanGroupPairs()
    {
        List<Group> groups = XmlReader.GetGroups();
        List<Plan> plans = new List<Plan>();
        List<PlanGroupPair> planGroupPairs = new List<PlanGroupPair>();
        foreach (var plxFilePath in _plxFilesPaths)
        {
            PlxReader plxReader = new PlxReader(plxFilePath);
            plans.Add(new Plan(plxFilePath, plxReader.GetPlanData()));
        }
        
        foreach (Plan plan in plans)
        {
            List<Group> matchingGroups = groups.Where(group => 
                group.Direction == plan.Data.Direction && 
                _selectedYear - plan.Data.YearOfAdmission + 1 == group.Years).ToList(); // TODO: Fix for selecting proper groups

            if (matchingGroups.Count <= 0)
            {
                Debug.WriteLine("-> для плана не нашлось ни одной группы\n" +
                                $"-> путь: [{plan.Path}]\n" +
                                $"-> направление: [{plan.Data.Direction}]\n");
                continue;
            }
            planGroupPairs.Add(new PlanGroupPair(plan, matchingGroups));
        }

        return planGroupPairs;
    }
}