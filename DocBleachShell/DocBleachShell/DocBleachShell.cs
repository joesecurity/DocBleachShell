// License: MIT
// Copyright: Joe Security
// Dependencies: - DocBleach https://github.com/docbleach
//				 - Log4Net https://logging.apache.org/log4net/
//				 - Ntfs Streams https://github.com/RichardD2/NTFS-Streams

using System;
using System.Collections;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;

using System.Threading;
using log4net;
using Microsoft.Win32;
using Trinet.Core.IO.Ntfs;

[assembly: log4net.Config.XmlConfigurator(Watch = false)]

namespace DocBleachShell
{
	
	/// <summary>
	/// This class implements a simple Wrapper around DocBleach (https://github.com/docbleach) by using the Windows Shell handlers. The idea is to call docbleach
	/// for each document before Office is opening it. That way documents are sanitized "bleached" automatically.
	/// </summary>
	class DocBleachShell
	{

		private static String ParentDirectory;
		private static String AssemblyFilePath;
		
		private static readonly ILog Logger = LogManager.GetLogger(typeof(DocBleachShell));
		
		[DllImport("kernel32")]
   		static extern bool AllocConsole();
   		
   		[DllImport( "kernel32", SetLastError = true )]
  		static extern bool AttachConsole( int dwProcessId );
		
 		[DllImport("kernel32.dll")]
    	static extern bool FreeConsole();  		
    	
    	[DllImport("User32.Dll", EntryPoint = "PostMessageA")]
       	static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);

        const int VK_RETURN = 0x0D;
       	const int WM_KEYDOWN = 0x100;
       	
       	private static readonly String VERSION = "0.0.2";
  		
		/// <summary>
		/// Main routine called for install, uninstall + to bleach documents.
		/// </summary>
		/// <param name="args"></param>
		public static void Main(string[] Args)
		{
			AssemblyFilePath = System.Reflection.Assembly.GetEntryAssembly().Location;
			ParentDirectory = Directory.GetParent(AssemblyFilePath).FullName;
			
			if(Args.Length == 1)
			{
				if (!AttachConsole(-1)) AllocConsole();
				
				Console.WriteLine("");
				
				Logger.Info("DocBleachShell v" + VERSION);
				Logger.Info("Copyright Joe Security");
				Logger.Info("www.joesecurity.org");
				Logger.Info("License: MIT");
				
				// Requires admin.
				if(Args[0].Equals("-install"))
				{
					if(Helper.IsUserAdministrator())
					{
						Install();
					} else
					{
						Logger.Error("Please install as Administrator");
					}
	
				// Requires admin.
				} if(Args[0].Equals("-uninstall"))
				{
					if(Helper.IsUserAdministrator())
					{
						RestoreBackup();
					} else
					{
						Logger.Error("Please uninstall as Administrator");
					}

				}
		
				closeConsole();
                
				
			} else if (Args.Length > 1)
			{
				try
				{
					Bleach(Args);
				} catch(Exception e)
				{
					Logger.Error("Error during bleaching", e);
				}
			}
		}
		
		/// <summary>
		/// Replaces all Shell\Open\command for Winword.exe, Excel.exe and Powerpnt.exe to point to our assembly. Before registry entries are modified a backup
		/// is made to \backup. Also the assembly paths of winword, excel and powerpnt are stored.
		/// </summary>
		private static void Install()
		{
			Logger.Info("Installing DocBleachShell");

			try
			{
				Directory.CreateDirectory(ParentDirectory + "\\backup");
			} catch(Exception)
			{
			}
			
			try
			{
				Directory.CreateDirectory(ParentDirectory + "\\backup\\docs");
			} catch(Exception)
			{
			}
			
			RegistryKey Root = Registry.ClassesRoot;
			
			// Over all classes
	        foreach (String V in Root.GetSubKeyNames())
	        {
	        	RegistryKey Command = Registry.ClassesRoot.OpenSubKey(V + "\\Shell\\Open\\command");
	        	
	        	if(Command != null)
	        	{
	        		try
	        		{
	        		
		        		String Cmdline = ((String)Command.GetValue("")).ToLower();
	
		        		// Check for office
		        		if(Cmdline.Contains("winword.exe") || Cmdline.Contains("excel.exe") || Cmdline.Contains("powerpnt.exe"))
		        		{
		        			Logger.Debug("Try to replace: " + V + " real cmdline is " + Cmdline);
		        			
		        			// Backup
		        			if(!File.Exists(ParentDirectory + "\\backup\\backup_" + V + ".reg"))
		        			{
		        				Helper.LaunchCmd("reg export HKEY_CLASSES_ROOT\\" + V + " " + ParentDirectory + "\\backup\\backup_" + V + ".reg /y");
		        			}
		        			
		        			// Delete the legacy command and DDEXEC
		        			Helper.LaunchCmd("reg delete HKCR\\" + V + "\\Shell\\Open\\command /v command /f");
		        			Helper.LaunchCmd("reg delete HKCR\\" + V + "\\Shell\\Open\\ddeexec /f");
		        			
		        			
		        			// @="\"C:\\Program Files\\Microsoft Office 15\\Root\\Office15\\WINWORD.EXE\" /n \"%1\" /o \"%u\""
		        			
		        			int i = Cmdline.ToLower().IndexOf(".exe");
		        			
		        			String App = "";
		        			String NewPath = "";
		        			
		        			// Construct new path calling our assembly
		        			if(Cmdline.Contains("winword.exe"))
		        			{
		        				NewPath = "\"" + AssemblyFilePath + "\" /w " + Cmdline.Substring(i+4).Trim('"');
		        				App = "word";
		        			} else if(Cmdline.Contains("excel.exe"))
		        			{
		        				NewPath = "\"" + AssemblyFilePath + "\" /e " + Cmdline.Substring(i+4).Trim('"');
		        				App = "excel";	
		        			} else if(Cmdline.Contains("powerpnt.exe"))
		        			{
		        				NewPath = "\"" + AssemblyFilePath + "\" /p " + Cmdline.Substring(i+4).Trim('"');
		        				App = "powerpnt";
		        			}
		        			
		        			if(NewPath.Length != 0)
		        			{
		        				// Special handling for DDE
		        				if(!NewPath.Contains("%1"))
		        				{
		        					NewPath += " \"%1\"";
		        					NewPath = NewPath.Replace("/dde", "");	
		        				}
		        				
		        				NewPath = NewPath.Replace("\"", "\\\"");
		        				
		        				// Overwrite command
		        				Helper.LaunchCmd("reg add HKCR\\" + V + "\\Shell\\Open\\command /ve /d \"" + NewPath + "\" /f");
		        				
		        				Logger.Info("Successfully installed for: " + V);
		        			}
		        			
		        			// Store real path
		        			if(!File.Exists(ParentDirectory + "\\" + App))
	        				{
		        				File.WriteAllText(ParentDirectory + "\\" + App, Cmdline.Substring(0, i+4).Trim('"'));
	        				}
		        		}
	        		
	        		} catch(Exception e)
	        		{
	        			Logger.Error("Unable to replace class: " + V, e);
	        		}
	        	}
	        }
		}
		
		/// <summary>
		/// Restore registry backups.
		/// </summary>
		private static void RestoreBackup()
		{
			Logger.Info("Uninstall DocBleachShell");
			
			foreach(String F in Directory.GetFiles("backup"))
	        {
				Helper.LaunchCmd("regedit.exe /S " + F);
	        }
			
			Logger.Info("Uninstalled");
		}
		
		
		/// <summary>
		/// Called e.g. by explorer.exe. Missing is to call docbleach and then Office with the bleached document.
		/// </summary>
		/// <param name="Args"></param>
		private static void Bleach(string[] Args)
		{
			String OfficePath = "";
	
			// Get office path
			if(Args[0].Equals("/w"))
			{
				OfficePath = File.ReadAllText(ParentDirectory + "\\word");
			} else if(Args[0].Equals("/p"))
			{
				OfficePath = File.ReadAllText(ParentDirectory + "\\powerpnt");
			} else if(Args[0].Equals("/e"))
			{
				OfficePath = File.ReadAllText(ParentDirectory + "\\excel");
			}
			
			if(OfficePath.Length != 0)
			{
				int l = -1;
				String AllArgs = "";
				
				// Find document path.
				for(int i = 1; i < Args.Length; ++i)
				{
					if(Args[i].Contains("\\"))
					{
						AllArgs += "\"" + Args[i] + "\" ";
						l = i;
					} else{
						AllArgs += Args[i] + " ";
					}	
				}
				
				if(l == -1)
				{
					Logger.Error("Unable to find document: " + AllArgs);
					return;
				}
				
				bool doBleach = false;
				
				try
				{
					
					// Check ADS Zone
					String OnlyBleachInternetFiles =  ConfigurationManager.AppSettings["OnlyBleachInternetFiles"];
					
					if(bool.Parse(OnlyBleachInternetFiles))
					{
						FileInfo F = new FileInfo(Args[l]);
						
						if (F.AlternateDataStreamExists("Zone.Identifier"))
						{
						    AlternateDataStreamInfo DataStream = F.GetAlternateDataStream("Zone.Identifier", FileMode.Open);
						    using (TextReader Reader = DataStream.OpenText())
						    {
						       	String ADS = Reader.ReadToEnd();
						       
						       	if(ADS.Contains("ZoneId=3") || ADS.Contains("ZoneId=4"))
								{
									doBleach = true;
								}
						    }
						} else
						{
							Logger.Debug("No Zone.Identifier found");
						}
					}
					
				} catch(Exception e)
				{
					Logger.Error("Unable to check ADS for " + Args[l], e);
				}
				
				if(doBleach)
				{
					new DocBleachWrapper().Bleach(Args[l], ParentDirectory);
				} else
				{
					Logger.Debug("Do not bleach: " + Args[l] + " not from internet");
				}
				
				// Call Office with bleached file.
				try
				{
					Process.Start(OfficePath, AllArgs);
				} catch(Exception e)
				{
					Logger.Error("Unable to start office: " + OfficePath + " " + AllArgs, e);
				}
			}
		}

		/// <summary>
		/// Close the console, sends a final {ENTER} to the parent.
		/// </summary>
		private static void closeConsole()
		{
			FreeConsole();
			
			try
			{
				// Get parent process
				int PID = Process.GetCurrentProcess().Id;
		        ManagementObjectSearcher Search = new ManagementObjectSearcher("root\\CIMV2", 
				                                                               "SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = " + PID);
		        var Results = Search.Get().GetEnumerator();
		        Results.MoveNext();
		        var QueryObj = Results.Current;
		        uint PPID = (uint)QueryObj["ParentProcessId"];
		        Process ParentProcess = Process.GetProcessById((int)PPID);
				PostMessage(ParentProcess.MainWindowHandle, WM_KEYDOWN, VK_RETURN, 0);
			} catch(Exception)
			{
			} 
		}	
	}
}