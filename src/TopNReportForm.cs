using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using SAS.Shared.AddIns;

namespace SASPress.CustomTasks.TopN
{
	/// <summary>
	/// The form for the Top N report task.
	/// </summary>
	public class TopNReportForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.CheckBox chkIncludeChart;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.ComboBox cmbReport;
		private System.Windows.Forms.ComboBox cmbMeasure;
		private System.Windows.Forms.ComboBox cmbStatistic;
		private System.Windows.Forms.ComboBox cmbCategory;
		private System.Windows.Forms.NumericUpDown nudN;
		private System.Windows.Forms.Label lblN;
		private System.Windows.Forms.Label lbReport;
		private System.Windows.Forms.Label lblMeasure;
		private System.Windows.Forms.Label lblStatistic;
		private System.Windows.Forms.Label lblCategory;
		private System.Windows.Forms.TextBox txtTitle;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtFootnote;

		#region Properties

		private TopNReport _model = null;
		/// <summary>
		/// A handle to the "model" class that holds the task state
		/// and provides access to the application services.
		/// </summary>
		public TopNReport Model
		{
			set { _model = value; }
			get { return _model; }
		}
		
		#endregion
		
		#region Constructor and Dispose
		public TopNReportForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}
		#endregion

		#region Initialize controls from data and model settings

		private void LoadSettingsFromModel()
		{
			// if there is a report column specified and it appears in the column list, select it
			if (Model.ReportColumn.Length>0 && cmbReport.FindStringExact(Model.ReportColumn)!=-1)
			{
				cmbReport.SelectedItem = Model.ReportColumn;
			}
			else cmbReport.SelectedItem = selectReportPrompt;

			// if there is a category column specified and it appears in the column list, select it
			if (Model.CategoryColumn.Length>0 && cmbCategory.FindStringExact(Model.CategoryColumn)!=-1)
			{
				cmbCategory.SelectedItem = Model.CategoryColumn;
			}
			else cmbCategory.SelectedItem = noCategory;

			// if there is a measure column specified and it appears in the column list, select it
			if (Model.MeasureColumn.Length>0 && cmbMeasure.FindStringExact(Model.MeasureColumn)!=-1)
			{
				cmbMeasure.SelectedItem = Model.MeasureColumn;
			}
			else cmbMeasure.SelectedItem = selectMeasurePrompt;

			cmbStatistic.SelectedItem = Model.Statistic.ToString();

			// top how many?
			nudN.Value = Model.N;

			// whether to include the chart
			chkIncludeChart.Checked = Model.IncludeChart;

			// custom title and footnote for the report
			txtTitle.Text = Model.Title;
			txtFootnote.Text = Model.Footnote;
				
		}

		string noCategory = "<No category>";
		string selectReportPrompt = "<Select a report column>";
		string selectMeasurePrompt="<Select a measure>";
		System.Collections.Hashtable hashFormats = new Hashtable();
		private void LoadColumns()
		{
			cmbReport.Items.Add(selectReportPrompt);
			cmbMeasure.Items.Add(selectMeasurePrompt);

			try
			{
				// to populate the comboboxes with the available columns,
				// we have to examine the active data source.
				ISASTaskData data = Model.Consumer.ActiveData;
				ISASTaskDataAccessor da = data.Accessor;
				// to do that, we need to "open" the data to get a peek at the column
				// information
				if (da.Open())
				{
					for (int i=0; i<da.ColumnCount; i++)
					{
						// first add the most likely categories -- character and date columns
						ISASTaskDataColumn ci = da.ColumnInfoByIndex(i);
						if (ci.Group == VariableGroup.Character || ci.Group == VariableGroup.Date)
						{
							cmbReport.Items.Add(ci.Name);
							cmbCategory.Items.Add(ci.Name);
						}
						else 
						{
							cmbMeasure.Items.Add(ci.Name);
							hashFormats.Add(ci.Name, ci.Format);
						}
					}

					// now add the rest of the numerics for Report and Category.  
					// These are less likely to make sense, but we don't want to 
					// shut the door on "creative" reports.
					for (int i=0; i<da.ColumnCount; i++)
					{
						ISASTaskDataColumn ci = da.ColumnInfoByIndex(i);
						if (ci.Group == VariableGroup.Numeric)
						{
							cmbReport.Items.Add(ci.Name);
							cmbCategory.Items.Add(ci.Name);
						}
					}

					cmbCategory.Items.Add(noCategory);

					da.Close();
				}
				else
				{
					// something went wrong in trying to read the data, so 
					// report an error message and end the form.
					string dataname = "UNKNOWN";
					if (Model.Consumer.ActiveData!=null)
						dataname = string.Format("{0}.{1}",Model.Consumer.ActiveData.Library,Model.Consumer.ActiveData.Member);
					MessageBox.Show(string.Format("ERROR: Could not read column information from data {0}.",dataname));
					this.Close();
				}
			}
			catch
			{
				// log exception
			}
		}

		private void LoadStatistics()
		{
			cmbStatistic.Items.Add(TopNReport.eStatistic.Sum.ToString());
			cmbStatistic.Items.Add(TopNReport.eStatistic.Average.ToString());
			cmbStatistic.Items.Add(TopNReport.eStatistic.Count.ToString());
		}

		private void CommitChangesToModel()
		{
			Model.N = Convert.ToInt32(nudN.Value);
			Model.IncludeChart = chkIncludeChart.Checked;
			Model.Statistic = (TopNReport.eStatistic)Enum.Parse(typeof(TopNReport.eStatistic),
				cmbStatistic.SelectedItem.ToString(),false);
			Model.ReportColumn = cmbReport.SelectedItem.ToString();
			Model.MeasureColumn = cmbMeasure.SelectedItem.ToString();
			// store away the format we gleaned from when we read the data properties
			Model.MeasureFormat = (string)hashFormats[Model.MeasureColumn];
			if (cmbCategory.SelectedItem.ToString()==noCategory)
				Model.CategoryColumn="";
			else 
				Model.CategoryColumn=cmbCategory.SelectedItem.ToString();
			Model.Title = txtTitle.Text;
			Model.Footnote = txtFootnote.Text;
		}

		#endregion

		#region Overrides - OnLoad, OnClosing

		// Load up the column values and current selections
		protected override void OnLoad(EventArgs e)
		{
			base.OnLoad (e);

			// make sure this control does not drop below 1.
			nudN.Minimum = 1;

			if (Model != null)
			{
				LoadColumns();
				LoadStatistics();
				LoadSettingsFromModel();
			}

			// add handlers for events triggered when any of the combobox selections change
			this.cmbMeasure.SelectedIndexChanged += new System.EventHandler(this.OnSelectionChanged);
			this.cmbCategory.SelectedIndexChanged += new System.EventHandler(this.OnSelectionChanged);
			this.cmbStatistic.SelectedIndexChanged += new System.EventHandler(this.OnSelectionChanged);
			this.cmbReport.SelectedIndexChanged += new System.EventHandler(this.OnSelectionChanged);

			// update the controls based on current selections
			UpdateControls();
		}

		// window is closing
		protected override void OnClosing(CancelEventArgs e)
		{
			// If the user clicked OK, call back into the application to make sure that
			// we can close the window.  The user might decide
			// to cancel the decision to close.
			// You know that window that pops up when you re-run a task
			// that says "Do you want to replace your results?  Yes, No, or Cancel?"
			// Well, this call back allows us to respect the "Cancel" selection
			// before the task window is actually closed.
			if (DialogResult == DialogResult.OK && !Model.Consumer.VerifyTaskClosing(this))
			{
				DialogResult = DialogResult.None;
				e.Cancel = true;
			}

			base.OnClosing (e);
		}

		protected override void OnClosed(EventArgs e)
		{
			if (DialogResult.OK == DialogResult)
				CommitChangesToModel();
			base.OnClosed (e);
		}

		#endregion

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(TopNReportForm));
			this.btnOK = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.lblN = new System.Windows.Forms.Label();
			this.lbReport = new System.Windows.Forms.Label();
			this.cmbReport = new System.Windows.Forms.ComboBox();
			this.lblMeasure = new System.Windows.Forms.Label();
			this.lblStatistic = new System.Windows.Forms.Label();
			this.lblCategory = new System.Windows.Forms.Label();
			this.cmbMeasure = new System.Windows.Forms.ComboBox();
			this.cmbStatistic = new System.Windows.Forms.ComboBox();
			this.cmbCategory = new System.Windows.Forms.ComboBox();
			this.nudN = new System.Windows.Forms.NumericUpDown();
			this.chkIncludeChart = new System.Windows.Forms.CheckBox();
			this.txtTitle = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtFootnote = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.nudN)).BeginInit();
			this.SuspendLayout();
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.btnOK.Location = new System.Drawing.Point(299, 394);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(90, 26);
			this.btnOK.TabIndex = 15;
			this.btnOK.Text = "OK";
			// 
			// btnCancel
			// 
			this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.btnCancel.Location = new System.Drawing.Point(404, 394);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(90, 26);
			this.btnCancel.TabIndex = 16;
			this.btnCancel.Text = "Cancel";
			// 
			// lblN
			// 
			this.lblN.Location = new System.Drawing.Point(7, 176);
			this.lblN.Name = "lblN";
			this.lblN.Size = new System.Drawing.Size(200, 23);
			this.lblN.TabIndex = 8;
			this.lblN.Text = "Include how many top values?";
			this.lblN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lbReport
			// 
			this.lbReport.Location = new System.Drawing.Point(7, 16);
			this.lbReport.Name = "lbReport";
			this.lbReport.Size = new System.Drawing.Size(200, 23);
			this.lbReport.TabIndex = 0;
			this.lbReport.Text = "Report on which column?";
			this.lbReport.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbReport
			// 
			this.cmbReport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cmbReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbReport.Location = new System.Drawing.Point(216, 16);
			this.cmbReport.Name = "cmbReport";
			this.cmbReport.Size = new System.Drawing.Size(280, 24);
			this.cmbReport.TabIndex = 1;
			// 
			// lblMeasure
			// 
			this.lblMeasure.Location = new System.Drawing.Point(7, 56);
			this.lblMeasure.Name = "lblMeasure";
			this.lblMeasure.Size = new System.Drawing.Size(200, 23);
			this.lblMeasure.TabIndex = 2;
			this.lblMeasure.Text = "Which column to measure?";
			this.lblMeasure.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblStatistic
			// 
			this.lblStatistic.Location = new System.Drawing.Point(7, 96);
			this.lblStatistic.Name = "lblStatistic";
			this.lblStatistic.Size = new System.Drawing.Size(200, 23);
			this.lblStatistic.TabIndex = 4;
			this.lblStatistic.Text = "Apply which statistic?";
			this.lblStatistic.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCategory
			// 
			this.lblCategory.Location = new System.Drawing.Point(7, 136);
			this.lblCategory.Name = "lblCategory";
			this.lblCategory.Size = new System.Drawing.Size(200, 23);
			this.lblCategory.TabIndex = 6;
			this.lblCategory.Text = "Report across category?";
			this.lblCategory.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbMeasure
			// 
			this.cmbMeasure.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cmbMeasure.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbMeasure.Location = new System.Drawing.Point(216, 56);
			this.cmbMeasure.Name = "cmbMeasure";
			this.cmbMeasure.Size = new System.Drawing.Size(280, 24);
			this.cmbMeasure.TabIndex = 3;
			// 
			// cmbStatistic
			// 
			this.cmbStatistic.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cmbStatistic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbStatistic.Location = new System.Drawing.Point(216, 96);
			this.cmbStatistic.Name = "cmbStatistic";
			this.cmbStatistic.Size = new System.Drawing.Size(280, 24);
			this.cmbStatistic.TabIndex = 5;
			// 
			// cmbCategory
			// 
			this.cmbCategory.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cmbCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbCategory.Location = new System.Drawing.Point(216, 136);
			this.cmbCategory.Name = "cmbCategory";
			this.cmbCategory.Size = new System.Drawing.Size(280, 24);
			this.cmbCategory.TabIndex = 7;
			// 
			// nudN
			// 
			this.nudN.Location = new System.Drawing.Point(216, 176);
			this.nudN.Name = "nudN";
			this.nudN.Size = new System.Drawing.Size(56, 22);
			this.nudN.TabIndex = 9;
			// 
			// chkIncludeChart
			// 
			this.chkIncludeChart.Location = new System.Drawing.Point(8, 224);
			this.chkIncludeChart.Name = "chkIncludeChart";
			this.chkIncludeChart.Size = new System.Drawing.Size(256, 24);
			this.chkIncludeChart.TabIndex = 10;
			this.chkIncludeChart.Text = "Include chart in report";
			// 
			// txtTitle
			// 
			this.txtTitle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtTitle.Location = new System.Drawing.Point(8, 296);
			this.txtTitle.Name = "txtTitle";
			this.txtTitle.Size = new System.Drawing.Size(488, 22);
			this.txtTitle.TabIndex = 12;
			this.txtTitle.Text = "";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 272);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(488, 23);
			this.label1.TabIndex = 11;
			this.label1.Text = "Enter a report title (optional):";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 328);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(488, 23);
			this.label2.TabIndex = 13;
			this.label2.Text = "Enter a report footnote (optional):";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFootnote
			// 
			this.txtFootnote.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFootnote.Location = new System.Drawing.Point(8, 352);
			this.txtFootnote.Name = "txtFootnote";
			this.txtFootnote.Size = new System.Drawing.Size(488, 22);
			this.txtFootnote.TabIndex = 14;
			this.txtFootnote.Text = "";
			// 
			// TopNReportForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(512, 438);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtFootnote);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtTitle);
			this.Controls.Add(this.chkIncludeChart);
			this.Controls.Add(this.nudN);
			this.Controls.Add(this.cmbCategory);
			this.Controls.Add(this.cmbStatistic);
			this.Controls.Add(this.cmbMeasure);
			this.Controls.Add(this.lblCategory);
			this.Controls.Add(this.lblStatistic);
			this.Controls.Add(this.lblMeasure);
			this.Controls.Add(this.cmbReport);
			this.Controls.Add(this.lbReport);
			this.Controls.Add(this.lblN);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnOK);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.MinimumSize = new System.Drawing.Size(520, 470);
			this.Name = "TopNReportForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "SAS Custom Task";
			((System.ComponentModel.ISupportInitialize)(this.nudN)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Event handlers

		// raised when a selection in one of the comboboxes has changed
		private void OnSelectionChanged(object sender, System.EventArgs e)
		{
			UpdateControls();
		}

		private void UpdateControls()
		{
			// if "Count" is the selected statistic, then the measure column does not apply
			bool isCountSelected = (cmbStatistic.SelectedItem.ToString() == TopNReport.eStatistic.Count.ToString());
			cmbMeasure.Enabled = !isCountSelected;

			// disable the OK button unless a report column and measure are selected
			if ((cmbReport.SelectedItem.ToString()==selectReportPrompt) || 
				(cmbMeasure.SelectedItem.ToString()==selectMeasurePrompt && !isCountSelected) )
				btnOK.Enabled = false;
			else btnOK.Enabled = true;
		}

		#endregion

	}
}
