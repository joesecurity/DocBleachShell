// License: MIT
// Copyright: Joe Security
// Dependencies: - DocBleach https://github.com/docbleach
//				 - Log4Net https://logging.apache.org/log4net/
//				 - Ntfs Streams https://github.com/RichardD2/NTFS-Streams

using System;
using System.Configuration;
using System.IO;
using System.Threading;
using log4net;

namespace DocBleachShell
{
	/// <summary>
	/// Wrapper around DocBleach.
	/// </summary>
	public class DocBleachWrapper
	{
		private static readonly ILog Logger = LogManager.GetLogger(typeof(DocBleachWrapper));
		
		/// <summary>
		/// Bleach the document.
		/// </summary>
		/// <param name="FilePath"></param>
		/// <param name="TargetDirectory"></param>
		public void Bleach(String FilePath, String TargetDirectory)
		{
			Logger.Debug("Try to bleach: " + FilePath);
			
			String CurrentDir = Directory.GetCurrentDirectory();
			
			String AppDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DocBleachShell";

			// Create backup dir.
			try
			{
				Directory.CreateDirectory(AppDir + "\\backup\\docs");
			} catch(Exception)
			{
			}
					
			Directory.SetCurrentDirectory(Directory.GetParent(FilePath).FullName);
		
	
			String MakeBackup =  ConfigurationManager.AppSettings["MakeBackup"];
					
			if(bool.Parse(MakeBackup))
			{
				
				try
				{
					File.Copy(FilePath, AppDir + "\\backup\\docs\\" + Path.GetFileName(FilePath), true);
				} catch(Exception e)
				{
					Logger.Error("Unable to make a backup of the document", e);
				}
				
			}
			
			// Original doc -> bleach. .... do
			
			String TmpDoc = Path.GetDirectoryName(FilePath) + "\\bleach." + Path.GetFileName(FilePath);
			
			try
			{
				File.Move(FilePath, TmpDoc);
			} catch(Exception e)
			{
				Logger.Error("Unable to rename document: " + FilePath, e);
			}
			
			// Call docbleach which will generate doc
			Helper.LaunchCmd("java -jar \"" + TargetDirectory + "\\docbleach.jar\" -in \"" + TmpDoc + 
			          "\" -out \"" + Path.GetFileName(FilePath) + "\" > \"" + AppDir + "\\tmp.log\" 2>&1");
			
			String Output = File.ReadAllText(AppDir + "\\tmp.log");
			
			Logger.Debug("DocBleach output: " + Output);
			
			// If the document was bleach "contains potential malicious elements" analyze the file with Joe Sandbox Cloud.
			if(!Output.Contains("file was already safe"))
			{
				String APIKey = ConfigurationManager.AppSettings["JoeSandboxCloudAPIKey"];
				
				if(APIKey.Length != 0)
				{
					new JoeSandboxClient().Analyze(TmpDoc, APIKey);
				}
			}
			else
			{
				Logger.Debug("Doc not sent to cloud : no API key configured");
			}
			// Cleanup & recovery
			if(File.Exists(FilePath))
			{
				try
				{
					File.Delete(TmpDoc);
				} catch(Exception e)
				{
					Logger.Error("Unable to delete original file " + TmpDoc, e);
				}
				Logger.Debug("Successfully bleached: " + FilePath);
			} else
			{
				Logger.Debug("Unable to bleach: " + FilePath);
				
				// No doc, move back.
				try
				{
					File.Move(TmpDoc, FilePath);
					
				} catch(Exception e)
				{
					Logger.Error("Unable to rename document: " + TmpDoc, e);
				}
			}
			
			Directory.SetCurrentDirectory(CurrentDir);

		}
	}
}
