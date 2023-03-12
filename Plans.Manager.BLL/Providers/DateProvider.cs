namespace Plans.Manager.BLL.Providers;

public static class DateProvider
{
	public static int CurrentStudyYear
	{
		get
		{
			DateTime dateTime = DateTime.Now;
			int currentYear = dateTime.Year;
			if(dateTime.Month <= 8)
				currentYear--;
			return currentYear;
		}
	}

	public static string CurrentStudyYearRange
	{
		get
		{
			DateTime dateTime = DateTime.Now;
			int currentYear = dateTime.Year;
			if(dateTime.Month <= 8)
				currentYear--;
			return $"{currentYear}-{currentYear + 1}";
		}
	}
}