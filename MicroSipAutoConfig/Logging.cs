using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;

namespace MicroSipAutoConfig {
	public class Logging {
		public static string LOG_FILE_NAME = @"C:\Temp\" + Assembly.GetExecutingAssembly().GetName().Name + "_" + Environment.UserName + ".log";

		public static void ToLog(string msg) {
			try {
				using (System.IO.StreamWriter sw = System.IO.File.AppendText(LOG_FILE_NAME)) {
					string logLine = string.Format("{0:G}: {1}", System.DateTime.Now, msg);
					sw.WriteLine(logLine);
				}
			} catch (Exception e) {
				Console.WriteLine("LogMessageToFile exception: " + LOG_FILE_NAME + Environment.NewLine + e.Message + 
					Environment.NewLine + e.StackTrace);
			}

			Console.WriteLine(msg);
		}
	}
}
