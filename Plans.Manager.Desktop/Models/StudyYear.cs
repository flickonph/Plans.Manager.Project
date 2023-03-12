using System;

namespace Plans.Manager.Desktop.Models;

public static class StudyYear
{
	public static int CurrentStudyYear()
	{
		DateTime date = DateTime.Now;
		int currentYear = date.Year;
		if(date.Month <= 8)
			currentYear--;
		return currentYear;
	}
}