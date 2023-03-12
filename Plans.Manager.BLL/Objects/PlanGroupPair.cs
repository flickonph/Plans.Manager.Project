namespace Plans.Manager.BLL.Objects;

public class PlanGroupPair
{
    public PlanGroupPair(Plan plan, List<Group> groups)
    {
        Plan = plan;
        Groups = groups;
    }
    public Plan Plan { get; }
    public List<Group> Groups { get; }
}