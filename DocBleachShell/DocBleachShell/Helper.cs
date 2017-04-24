// License: MIT
// Copyright: Joe Security
// Dependencies: - DocBleach https://github.com/docbleach
//				 - Log4Net https://logging.apache.org/log4net/
//				 - Ntfs Streams https://github.com/RichardD2/NTFS-Streams

using System;
using System.Diagnostics;
using System.Security.Principal;
using log4net;

namespace DocBleachShell
{
	/// <summary>
	/// Helper collection.
	/// </summary>
	public class Helper
	{
		private static readonly ILog Logger = LogManager.GetLogger(typeof(Helper));
			
		/// <summary>
		/// Helper to call command line (hidden).
		/// </summary>
		/// <param name="cmd"></param>
		public static void LaunchCmd(String cmd)
		{
			try
			{
				ProcessStartInfo StartInfo = new ProcessStartInfo();
				
				StartInfo.FileName = "cmd.exe";
				StartInfo.Arguments = "/C " + cmd;
				StartInfo.RedirectStandardOutput = true;
				StartInfo.RedirectStandardError = true;
				StartInfo.UseShellExecute = false;
				StartInfo.CreateNoWindow = true;
				
				Process Proc = new Process();
				
				Proc.StartInfo = StartInfo;
				Proc.EnableRaisingEvents = true;
				Proc.Start();
				Proc.WaitForExit();
				
			} catch(Exception e)
			{
				Logger.Error("Unable to start cmd: " + cmd, e);
			}
		}
		
		/// <summary>
		/// Helper to check admin status.
		/// </summary>
		/// <returns></returns>
		public static bool IsUserAdministrator()
		{
		    bool IsAdmin;
		    
		    try
		    {
		        WindowsIdentity User = WindowsIdentity.GetCurrent();
		        WindowsPrincipal Principal = new WindowsPrincipal(User);
		        IsAdmin = Principal.IsInRole(WindowsBuiltInRole.Administrator);
		    }
		    catch (UnauthorizedAccessException)
		    {
		        IsAdmin = false;
		    }
		    catch (Exception)
		    {
		        IsAdmin = false;
		    }
		    
		    return IsAdmin;
		}     
		
	}
}
