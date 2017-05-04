// License: MIT
// Copyright: Joe Security
// Dependencies: - DocBleach https://github.com/docbleach
//				 - Log4Net https://logging.apache.org/log4net/
//				 - Ntfs Streams https://github.com/RichardD2/NTFS-Streams


using System;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using log4net;

namespace DocBleachShell
{
	/// <summary>
	/// Minimalistic JoeSandboxClient. Only implements upload functionality.
	/// </summary>
	public class JoeSandboxClient
	{
		private static readonly ILog Logger = LogManager.GetLogger(typeof(JoeSandboxClient));
		
		private const String APIUrl = "https://jbxcloud.joesecurity.org/api/analysis";
		
				/// <summary>
		/// Analyze the file on Joe Sandbox Cloud. 
		/// </summary>
		/// <param name="FilePath"></param>
		/// <param name="APIKey"></param>
		public void Analyze(String FilePath, String APIKey)
		{
			
			Logger.Debug("Analyze: " + FilePath + " with Joe Sandbox Cloud");
			
			try {
				
				WebRequest Request = WebRequest.Create(APIUrl);

				string Boundary = "----------------------------" + DateTime.Now.Ticks.ToString("x");
				
				Request.Method = "POST";
				Request.Timeout = 15000;
				Request.ContentType = "multipart/form-data; boundary=" + Boundary;

				NameValueCollection FormData = new NameValueCollection();

				FormData["apikey"] = APIKey;
				
				// Allow full interenet activity.
				FormData["inet"] = "1";
				
				// Hybrid Code Analysis = true.
				FormData["scae"] = "1";
				
				// Enable Hybrid Decompilation.
				FormData["dec"] = "1";
				
				// Enable VBA Macro instrumentation.
				FormData["vbainstr"] = "1";
				
				FormData["tandc"] = "1";
				FormData["type"] = "file";
				
				// Auto system selection.
				FormData["auto"] = "1";
				
				FormData["comments"] = "Submitted by DocBleachShell";

				Stream DataStream = getPostStream(FilePath, FormData, Boundary);

				Request.ContentLength = DataStream.Length;

				Stream ReqStream = Request.GetRequestStream();

				DataStream.Position = 0;

				byte[] Buffer = new byte[1024];
				int BytesRead = 0;

				while ((BytesRead = DataStream.Read(Buffer, 0, Buffer.Length)) != 0) {
					ReqStream.Write(Buffer, 0, BytesRead);
				}

				DataStream.Close();
				ReqStream.Close();

				StreamReader Reader = new StreamReader(Request.GetResponse().GetResponseStream());
				Logger.Debug("Joe Sandbox Cloud answer: " + Reader.ReadToEnd());
				Logger.Debug("Successfully submit file to Joe Sandbox Cloud");

			} catch (Exception e) {
				Logger.Error("Unable to analyze file: " + FilePath + " with Joe Sandbox", e);

			}
		}
		
		/// <summary>
		/// Building form and file data. 
		/// </summary>
		/// <param name="FilePath"></param>
		/// <param name="FormData"></param>
		/// <param name="Boundary"></param>
		/// <returns></returns>
		private Stream getPostStream(string FilePath, NameValueCollection FormData, string Boundary)
		{
			Stream PostDataStream = new System.IO.MemoryStream();

			// Adding form data.
			string FormDataHeaderTemplate = Environment.NewLine + "--" + Boundary + Environment.NewLine +
				"Content-Disposition: form-data; name=\"{0}\";" + Environment.NewLine + Environment.NewLine + "{1}";

			foreach (string Key in FormData.Keys) {
				byte[] FormItemBytes = System.Text.Encoding.UTF8.GetBytes(string.Format(FormDataHeaderTemplate,
				                                                                        Key, FormData[Key]));
				PostDataStream.Write(FormItemBytes, 0, FormItemBytes.Length);
			}

			if (FilePath != null) {

				// Adding file data. 
				FileInfo fileInfo = new FileInfo(FilePath);
				
				string FileHeaderTemplate = Environment.NewLine + "--" + Boundary + Environment.NewLine +
					"Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"" +
					Environment.NewLine + "Content-Type: application/octet-stream; " + Environment.NewLine + Environment.NewLine;

				byte[] FileHeaderBytes = System.Text.Encoding.UTF8.GetBytes(string.Format(FileHeaderTemplate,
				                                                                          "sample", fileInfo.FullName));
				PostDataStream.Write(FileHeaderBytes, 0, FileHeaderBytes.Length);

				FileStream Stream = fileInfo.OpenRead();

				// Write form + file. 
				byte[] Buffer = new byte[1024];
				int BytesRead = 0;

				while ((BytesRead = Stream.Read(Buffer, 0, Buffer.Length)) != 0) {
					PostDataStream.Write(Buffer, 0, BytesRead);
				}

				Stream.Close();
			}

			// Ending.
			byte[] EndBoundaryBytes = System.Text.Encoding.UTF8.GetBytes(Environment.NewLine + "--" + Boundary + "--");
			PostDataStream.Write(EndBoundaryBytes, 0, EndBoundaryBytes.Length);

			return PostDataStream;
		}
	}
}
