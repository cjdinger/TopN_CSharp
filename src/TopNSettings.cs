using System;

namespace SASPress.CustomTasks.TopN
{
	/// <summary>
	/// A serializable class that represents the settings for the Top N report.
	/// </summary>
	[Serializable]
	public class TopNSettings
	{
		/// <summary>
		/// a member variable that tracks the report column name
		/// </summary>
		private string _reportColumn="";
		/// <summary>
		/// The property used by other classes to set/get the report column name.
		/// </summary>
		public string ReportColumn
		{
			get { return _reportColumn; }
			set { _reportColumn = value; }
		}
		private string _categoryColumn="";
		/// <summary>
		/// The property used by other classes to set/get the category column name.
		/// </summary>
		public string CategoryColumn
		{
			get { return _categoryColumn; }
			set { _categoryColumn = value; }
		}

		private string _measureColumn = "";
		/// <summary>
		/// The property used by other classes to set/get the measure column name.
		/// </summary>
		public string MeasureColumn
		{
			get { return _measureColumn; }
			set { _measureColumn = value; }
		}

		private bool _includeChart = true;
		/// <summary>
		/// The property used by other classes to get/set whether to 
		/// include a chart in the report.
		/// </summary>
		public bool IncludeChart 
		{
			get { return _includeChart; }
			set { _includeChart = value; }
		}

		private int _n=10;
		/// <summary>
		/// The property used by other classes to set/get the
		/// value of "N", and important component of a "Top N" report,
		/// wouldn't you agree?
		/// </summary>
		/// <exception cref="System.ArgumentOutOfRangeException">
		/// Thrown if the value is set to something less than 1.
		/// The "Top minus 10" report doesn't really make sense, does it?
		/// </exception>
		public int N
		{
			get { return _n; }
			set 
			{
				if (_n>0)
					_n = value; 
				else 
					throw new ArgumentOutOfRangeException("The value for N must be greater than zero.");
			}
		}

		private TopNReport.eStatistic _statistic = TopNReport.eStatistic.Sum;
		/// <summary>
		/// The statistic that we are measuring: an enum representing Sum, Average, or frequency Count.
		/// </summary>
		public TopNReport.eStatistic Statistic
		{
			get { return _statistic; }
			set { _statistic = value; }
		}

		private string _title = "";
		/// <summary>
		/// Custom title for the report
		/// </summary>
		public string Title
		{
			get { return _title; }
			set { _title = value; }
		}

		private string _footnote = "";
		/// <summary>
		/// Custom footnote for the report
		/// </summary>
		public string Footnote
		{
			get { return _footnote; }
			set { _footnote = value; }
		}

		private string _measureFormat = "";
		/// <summary>
		/// Does the measure column have a format (such as currency)?  If so,
		/// this field will help to propogate it to the report.
		/// </summary>
		public string MeasureFormat
		{
			get { return _measureFormat; }
			set { _measureFormat = value; }
		}
	}
}
