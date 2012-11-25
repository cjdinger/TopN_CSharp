using System;
using System.Xml;		// for XMLTextWriter and XMLTextReader
using System.IO;		// for StreamReader and StreamWriter
using System.Text;		// for StringBuilder
using System.Xml.Serialization; // for serialization

// SAS AddIns namespace
using SAS.Shared.AddIns;

namespace SASPress.CustomTasks.TopN
{
	/// <summary>
	/// Top N task for use in SAS Enterprise Guide and SAS Add-In for Microsoft Office
	/// </summary>
	public class TopNReport : 
		SAS.Shared.AddIns.ISASTaskAddIn, 
		SAS.Shared.AddIns.ISASTaskDescription, 
		SAS.Shared.AddIns.ISASTask
	{
		#region private members for boilerplate values
		
		// These are the labels/descriptions that show up in the
		// task menus within the application.
		private string sLabel = "Top N Report";
		private string sAddInName = "Top N Report";
		private string sAddInDescription = "Creates a report and chart of the top values within your data.";
		private string sProductsReq = "BASE; GRAPH";
		private string sProductsOpt = "";
		private string sTaskName= "&Top N Report";
		private string sCategory = "SAS Press Examples";
		private string sTaskDescription = "Creates a report and chart of the top values within your data.";
		private string sFriendlyName = "Top N Report";
		private string sWhatIsDescription = "Creates a report and chart of the top values within your data.";
		
		#endregion

		#region Properties for task state
		// DESIGN TIP:
		// Each property has a member variable that keeps track of the value
		// and a public "setter/getter" that allows other classes
		// to get and set that value.
		// Does it seem like a silly exercise?  Why not simply make the 
		// member variable public, so other classes can set/get its value
		// directly?  We could do that, but that isn't considered
		// good object-oriented design.  The fact is, how we keep track
		// of property values in this class is a "private" matter.  
		// By surfacing getter/setter properties, we are free to change
		// the implementation of how we track the value without affecting 
		// the other classes that might rely on it.
		// Also, we can implement some basic validation in the setter.
		// See the implementation of the "N" property for an example.

		/// <summary>
		/// The possible values for the statistic used in the report
		/// </summary>
		public enum eStatistic
		{
			Sum,
			Average,
			Count
		}

		/// <summary>
		/// a member variable that tracks the report column name
		/// </summary>
		private string _reportColumn="";
		/// <summary>
		/// The property used by other classes to set/get the report column name.
		/// </summary>
		internal string ReportColumn
		{
			get { return _reportColumn; }
			set { _reportColumn = value; }
		}
		private string _categoryColumn="";
		/// <summary>
		/// The property used by other classes to set/get the category column name.
		/// </summary>
		internal string CategoryColumn
		{
			get { return _categoryColumn; }
			set { _categoryColumn = value; }
		}

		private string _measureColumn = "";
		/// <summary>
		/// The property used by other classes to set/get the measure column name.
		/// </summary>
		internal string MeasureColumn
		{
			get { return _measureColumn; }
			set { _measureColumn = value; }
		}

		private string _measureFormat = "";
		/// <summary>
		/// The value for the SAS Format that we should apply to the measure column
		/// </summary>
		internal string MeasureFormat
		{
			get { return _measureFormat; }
			set { _measureFormat = value; }
		}

		private bool _includeChart = true;
		/// <summary>
		/// The property used by other classes to get/set whether to 
		/// include a chart in the report.
		/// </summary>
		internal bool IncludeChart 
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
		internal int N
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

		private eStatistic _statistic = eStatistic.Sum;
		internal eStatistic Statistic
		{
			get { return _statistic; }
			set { _statistic = value; }
		}


		private string _title = "";
		internal string Title
		{
			get { return _title; }
			set { _title = value; }
		}

		private string _footnote = "";
		internal string Footnote
		{
			get { return _footnote; }
			set { _footnote = value; }
		}
		#endregion

		// store reference to the application Consumer object
		private ISASTaskConsumer consumer = null;

		/// <summary>
		/// Property getter for the ISASTaskConsumer handle
		/// </summary>
		public ISASTaskConsumer Consumer
		{
			get { return consumer; }
		}

		public TopNReport()
		{
		}

		#region ISASTaskAddIn Members

		public bool VisibleInManager
		{
			get
			{				
				return true;
			}
		}

		public string AddInName
		{
			get
			{				
				return sAddInName;
			}
		}

		public void Disconnect()
		{
			// perform cleanup tasks
			consumer = null;
		}

		public int Languages(out string[] Items)
		{		
			// by default, we support English 
			// Add more languages as needed
			Items = new string[] {"en-US"};
			return 1;
		}

		public string AddInDescription
		{
			get
			{				
				return sAddInDescription;
			}
		}

		public bool Connect(ISASTaskConsumer Consumer)
		{
			// perform any initialization needed when the application connects
			consumer = Consumer;
	
			// this is a good time to get the consumer.ActiveData, if your task requires it.

			return true;
		}

		public string Language
		{
			set
			{
				// if you support multiple languages, add handler here
			}
		}

		#endregion

		#region ISASTaskDescription Members

		public string ProductsRequired
		{
			get
			{
				// What SAS products are required for this task to run?
				return sProductsReq;
			}
		}

		public bool StandardCategory
		{
			get
			{
				// used typically only by SAS-supplied tasks			
				return false;
			}
		}

		public bool GeneratesListOutput
		{
			get
			{
				// Does this task generate ODS-style output?
				return true;
			}
		}

		public string TaskName
		{
			get
			{
				return sTaskName;
			}
		}

		public string Validation
		{
			get
			{
				// Add a validation string that makes sense for your organization
				// For example, "dev", "test", or "prod"
				return "";
			}
		}

		public string TaskCategory
		{
			get
			{
				// PLACEHOLDER: replace with your own category
				return sCategory;
			}
		}

		public string IconAssembly
		{
			get
			{
				// return the full path/name of this assembly, assuming that the
				// icon is embedded within the assembly
				return System.Reflection.Assembly.GetExecutingAssembly().Location;
			}
		}

		public string Clsid
		{
			get
			{
				// ClassID GUID generated by the template
				return "7EFFA4BF-2E74-40C9-92B4-8841545909AE";
			}
		}

		public SAS.Shared.AddIns.ShowType TaskType
		{
			get
			{				
				return SAS.Shared.AddIns.ShowType.Wizard;
			}
		}

		public bool RequiresData
		{
			get
			{
				// Does your task require input data from the application?
				return true;
			}
		}

		public string IconName
		{
			get
			{
				// return the name of the icon within this assembly
				// including namespace qualifiers
				return "SASPress.CustomTasks.TopN.TopN.ico";
			}
		}

		public int MinorVersion
		{
			get
			{
				return 0;
			}
		}

		public int MajorVersion
		{
			get
			{
				return 1;
			}
		}

		public int NumericColsRequired
		{
			get
			{
				// How many numeric variables are required in input data, if any?
				return 0;
			}
		}

		public string TaskDescription
		{
			get
			{				
				return sTaskDescription;
			}
		}

		public bool GeneratesSasCode
		{
			get
			{
				// Return true if your task generates SAS program code
				return true;
			}
		}

		public string FriendlyName
		{
			get
			{
				// Replace with a user-friendly name for your task
				return sFriendlyName;
			}
		}

		public string ProcsUsed
		{
			get
			{
				// What SAS procedures are used in this task?
				return "";
			}
		}

		public string WhatIsDescription
		{
			get
			{
				// Longer description for your task
				return sWhatIsDescription;
			}
		}

		public string ProductsOptional
		{
			get
			{
				// What SAS products are optionally used by this task?
				return sProductsOpt;
			}
		}

		#endregion

		#region ISASTask Members

		public string RunLog
		{
			get
			{
				// if your task does not generate SAS code, you can supply your own
				// log text to record the work completed.
				return "";
			}
		}

		// manage the state of the task when serializing to and from the project
		// makes use of some helper methods: WriteXML and ReadXML
		public string XmlState
		{
			get
			{
				return WriteXML();
			}
			set
			{
				ReadXML(value);
			}
		}

		public void Terminate()
		{
			// Cleanup as needed
		}

		public SAS.Shared.AddIns.OutputData OutputDataInfo(int Index, out string Source, out string Label)
		{
			// no output data created
			Source = null;
			Label = null;
			return SAS.Shared.AddIns.OutputData.Unknown;

		}

		const string countStep =	"data work._tpnview / view=work._tpnview; \n" + 
									"  set &data; \n" +
									"  _tpncount=1; \n" +
									"  label _tpncount='Count'; \n" +
									"run; \n" +
									"%let data=work._tpnview; \n";

		public string SasCode
		{
			get
			{	
				StringBuilder code = new StringBuilder();
				code.AppendFormat("%let data={0}.{1};\n",Consumer.ActiveData.Library, consumer.ActiveData.Member);
				code.AppendFormat("%let report={0};\n",_reportColumn);
				code.AppendFormat("%let measure={0};\n",_statistic==eStatistic.Count ? "_tpncount" : _measureColumn);
				// determine whether to apply a format appropriate for the measure
				string format="";
				if (_statistic!=eStatistic.Count && _measureFormat.Length>0)
					format = string.Format("%str(format={0})",_measureFormat);
				code.AppendFormat("%let measureformat={0};\n",format);
				code.AppendFormat("%let stat={0};\n",_statistic==eStatistic.Average ? "MEAN" : "SUM");
				code.AppendFormat("%let n={0};\n",_n.ToString(System.Globalization.CultureInfo.InvariantCulture));
				if (_title.Length>0)
					code.AppendFormat("title \"{0}\";\n",_title);
				else code.Append("title;\n");
				if (_footnote.Length>0) 
					code.AppendFormat("footnote \"{0}\";\n",_footnote);
				else code.Append("footnote;\n");
				
				// put in the data step to add a count column
				if (_statistic==eStatistic.Count)
					code.AppendFormat(countStep);

				// Read the code file from an embedded SAS file
				// If there are any substitutions or parameters to set for the SAS program,
				// modify the string accordingly before returning it.
				if (_categoryColumn.Length==0)
				{
					code.Append(ReadFileFromAssembly("SASPress.CustomTasks.TopN.Programs.StraightReport.sas"));
					if (_includeChart)
						code.Append(ReadFileFromAssembly("SASPress.CustomTasks.TopN.Programs.StraightChart.sas"));
				}
				else
				{
					code.AppendFormat("%let category={0};\n",_categoryColumn);
					code.Append(ReadFileFromAssembly("SASPress.CustomTasks.TopN.Programs.StratifiedReport.sas"));
					if (_includeChart)
						code.Append(ReadFileFromAssembly("SASPress.CustomTasks.TopN.Programs.StratifiedChart.sas"));

				}
				return code.ToString();
			}
		}

		public bool Initialize()
		{
			// initialize this instance of your task
			return true;
		}

		public SAS.Shared.AddIns.ShowResult Show(System.Windows.Forms.IWin32Window Owner)
		{
			// Show the default form for this custom task
			TopNReportForm dlg = new TopNReportForm();
			dlg.Model = this;
			dlg.Text = sLabel;

			if (dlg.ShowDialog()==System.Windows.Forms.DialogResult.OK)
				return SAS.Shared.AddIns.ShowResult.RunNow;
			else
				return SAS.Shared.AddIns.ShowResult.Canceled;
		}

		// The label this task is known by within a project
		public string Label
		{
			get
			{			
				return sLabel;
			}
			set
			{
				sLabel = value;
			}
		}

		public int OutputDataCount
		{
			get
			{				
				// This task does not create an output data set, so return 0.
				return 0;
			}
		}

		#endregion

		#region Helper methods for serialization 
		// Save the task state so that settings are remembered
		private string WriteXML()
		{
			TopNSettings settings = new TopNSettings();
			settings.CategoryColumn = CategoryColumn;
			settings.IncludeChart = IncludeChart;
			settings.MeasureColumn = MeasureColumn;
			settings.N = N;
			settings.ReportColumn = ReportColumn;
			settings.Statistic = Statistic;
			settings.Title = Title;
			settings.Footnote = Footnote;
			settings.MeasureFormat = MeasureFormat;

			using (StringWriter sw = new StringWriter())
			{
				XmlSerializer s = new XmlSerializer(typeof(TopNSettings));
				s.Serialize(sw, settings);
				return sw.ToString();
			}   

		}

		private void ReadXML(string xml)
		{
			if (xml!=null && xml.Length>0)
			{
				TopNSettings settings;

				try
				{
					using (StringReader sr = new StringReader(xml))
					{
						XmlSerializer s = new XmlSerializer(typeof(TopNSettings));
						settings = (TopNSettings)s.Deserialize(sr);
						CategoryColumn = settings.CategoryColumn;
						IncludeChart = settings.IncludeChart;
						MeasureColumn = settings.MeasureColumn;
						N = settings.N;
						ReportColumn = settings.ReportColumn;
						Statistic = settings.Statistic;
						Title = settings.Title;
						Footnote = settings.Footnote;
						MeasureFormat = settings.MeasureFormat;
					}				
				}
				catch
				{
				}
			}
		}

		#endregion

		#region utility methods
		/// <summary>
		/// Read an embedded text file out of the assembly and return it in a string
		/// </summary>
		/// <param name="filename">Name of the file to read, with namespace qualifiers</param>
		/// <returns>a string with the contents of the file</returns>
		internal static string ReadFileFromAssembly(string filename) 
		{
			string filecontents = String.Empty;
			System.Reflection.Assembly assem = System.Reflection.Assembly.GetCallingAssembly();
			Stream stream = assem.GetManifestResourceStream(filename);
			if (stream != null) 
			{
				StreamReader sr = new StreamReader(stream);
				filecontents = sr.ReadToEnd();
				sr.Close();
				stream.Close();
			}

			return filecontents;
		}
		#endregion
	}
}
