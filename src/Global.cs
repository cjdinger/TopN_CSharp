// ---------------------------------------------------------------
// Copyright 2005, SAS Institute Inc.
// ---------------------------------------------------------------
using System;
using System.Reflection;
using System.IO;
using System.Drawing;


namespace SASPress.CustomTasks.TopN
{
	/// <summary>
	/// central class for utility routines
	/// </summary>
	internal class Global
	{
		/// <summary>
		/// Read an embedded text file out of the assembly and return it in a string
		/// </summary>
		/// <param name="filename">Name of the file to read, with namespace qualifiers</param>
		/// <returns>a string with the contents of the file</returns>
		internal static string ReadFileFromAssembly(string filename) 
		{
			string filecontents = String.Empty;
			Assembly assem = System.Reflection.Assembly.GetCallingAssembly();
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
	}
}
