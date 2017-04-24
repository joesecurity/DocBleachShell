// License: MIT
// Copyright: Joe Security
// Dependencies: - DocBleach https://github.com/docbleach
//				 - Log4Net https://logging.apache.org/log4net/
//				 - Ntfs Streams https://github.com/RichardD2/NTFS-Streams

using System;
using System.Configuration;
using System.IO;
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
					
			Directory.SetCurrentDirectory(Directory.GetParent(FilePath).FullName);
		
			try
			{
				File.Copy(FilePath, TargetDirectory + "\\backup\\docs\\" + Path.GetFileName(FilePath), true);
			} catch(Exception e)
			{
				Logger.Error("Unable to make a backup of the document", e);
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
			          "\" -out \"" + Path.GetFileName(FilePath) + "\" > \"" + TargetDirectory + "\\tmp.log\" 2>&1");
			
			String Output = File.ReadAllText(TargetDirectory + "\\tmp.log");
			
			Logger.Debug("DocBleach output: " + Output);
			
			// If the document was bleach "contains potential malicious elements" analyze the file with Joe Sandbox Cloud.
			if(!Output.Contains("file was already safe"))
			{
				String APIKey = ConfigurationManager.AppSettings["JoeSandboxCloudAPIKey"];
				new JoeSandboxClient().Analyze(TmpDoc, APIKey);
			}
			
			// Cleanup & recovery
			if(File.Exists(FilePath))
			{
				File.Delete(TmpDoc);
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
