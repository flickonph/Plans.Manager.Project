using System;

namespace Plans.Manager.Desktop.Models;

public static class StudyYear
{
	public static int CurrentStudyYear()
	{
		var date = DateTime.Now;
		var currentYear = date.Year;
		if(date.Month <= 8)
			currentYear--;
		return currentYear;
	}
}