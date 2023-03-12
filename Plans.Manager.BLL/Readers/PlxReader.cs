using System.Text.RegularExpressions;
using System.Xml;
using Plans.Manager.BLL.Objects;

namespace Plans.Manager.BLL.Readers;

public class PlxReader : AReader
{
    public PlxReader(string path) : base(path) { }

    public PlanData GetPlanData()
    {
        // Public strings region
        string planDirectionCode = GetXmlElements("ООП").First().GetAttribute("Шифр");
        string planDirectionName = GetPlanDirection();
        string tempYearOfAdmission = GetXmlElements("Планы").First().GetAttribute("УчебныйГод");
        int planYearOfAdmission = Convert.ToInt32(tempYearOfAdmission.Length > 4 ? tempYearOfAdmission.Substring(0, 4) : tempYearOfAdmission);

        // Disciplines region
        List<Discipline> planDisciplines = new List<Discipline>();
        IEnumerable<XmlElement> defaultDisciplinesXmlElements = GetXmlElements("ПланыСтроки");
        foreach (XmlElement defaultDisciplineXmlElement in defaultDisciplinesXmlElements)
        {
            string disciplineName = defaultDisciplineXmlElement.GetAttribute("Дисциплина");
            string disciplineCode = defaultDisciplineXmlElement.GetAttribute("Код");
            string parentCode = defaultDisciplineXmlElement.GetAttribute("КодРодителя");
            string departmentCode = defaultDisciplineXmlElement.GetAttribute("КодКафедры");
            Department department = new Department(departmentCode, XmlReader.GetDepartmentName(departmentCode));

            List<DisciplineData> disciplineData = new List<DisciplineData>();
            IEnumerable<XmlElement> additionalDisciplineXmlElements = GetXmlElements("ПланыНовыеЧасы").Where(element => element.GetAttribute("КодОбъекта") == disciplineCode);
            foreach (XmlElement additionalDisciplineXmlElement in additionalDisciplineXmlElements)
            {
                string type = additionalDisciplineXmlElement.GetAttribute("КодВидаРаботы");
                double hours = Convert.ToDouble(additionalDisciplineXmlElement.GetAttribute("Количество"));
                int course = Convert.ToInt32(additionalDisciplineXmlElement.GetAttribute("Курс"));
                int semester = Convert.ToInt32(additionalDisciplineXmlElement.GetAttribute("Семестр"));
                /*var session = additionalDisciplineXmlElement.GetAttribute("Сессия");
                var weeks = additionalDisciplineXmlElement.GetAttribute("Недель");
                var days = additionalDisciplineXmlElement.GetAttribute("Дней");
                var hoursTypeCode = additionalDisciplineXmlElement.GetAttribute("КодТипаЧасов");*/

                DisciplineData disciplineDataItem = new DisciplineData(type, hours, semester, course);
                IEnumerable<XmlElement> typesXmlElements = GetXmlElements("СправочникВидыРабот");
                foreach (XmlElement element in typesXmlElements)
                    if (element.GetAttribute("Код") == type)
                        disciplineDataItem.Type = element.GetAttribute("Название");

                disciplineData.Add(disciplineDataItem);
            }

            planDisciplines.Add(new Discipline(disciplineName, disciplineCode, parentCode, disciplineData, department));
        }

        PlanData planData = new PlanData(planDirectionCode, planDirectionName, planYearOfAdmission, planDisciplines);

        return planData;
    }

    private string GetPlanDirection()
    {
        string direction;
        string fullString = GetXmlElements("Планы").First().GetAttribute("Титул");
        Regex regex = new Regex("(\")([^\"]*)(\"[^\"]*)");
        try
        {
            direction = regex.Match(fullString).Groups[2].Value;
            direction = direction.Replace('ё', 'е');
            direction = direction.Trim();
        }
        catch
        {
            direction = string.Empty;
        }

        return direction;
    }
}