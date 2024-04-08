using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Linq;
using System.IO.Pipes;
using DocumentFormat.OpenXml.Packaging;

namespace Descriptions
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string mainPath = "C:\\Users\\david.gomes\\Desktop\\descrições";

            string openDesc = "<description>";
            string closeDesc = "</description>";

            string openPt = "<description_pt>";
            string closePt = "</description_pt>";

            string openEn = "<description_en>";
            string closeEn = "</description_en>";

            string openFr = "<description_fr>";
            string closeFr = "</description_fr>";

            string openProperty = "<property>";
            string closeProperty = "</property>";

            string finalDoc = "";
            finalDoc += openDesc;

            string[] folders = CheckFoldersInside(mainPath);

            if (folders.Length > 0)
            {
                foreach (string folder in folders)
                {
                    string[] files = CheckFilesInside(folder);

                    string reference = "<reference>" + Path.GetFileName(folder) + "</reference>";
                    finalDoc += openProperty + reference;

                    foreach (string file in files)
                    {
                        string text = "";

                        if (file.Contains("PT"))
                        {
                            text = ReadWordDocument(file);
                            finalDoc += openPt + text + closePt;
                        }

                        if (file.Contains("ING"))
                        {
                            text = ReadWordDocument(file);
                            finalDoc += openEn + text + closeEn;
                        }

                        if (file.Contains("FR"))
                        {
                            text = ReadWordDocument(file);
                            finalDoc += openFr + text + closeFr;
                        }
                    }

                    finalDoc += closeProperty;
                }

                finalDoc += closeDesc;

                File.WriteAllText("C:\\Users\\david.gomes\\Desktop\\descrições_myro.txt", finalDoc);

                Console.WriteLine("XML file has been created successfully.");
            }

        }

        static string ReadWordDocument(string filePath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                return body.InnerText;
            }
        }

        static string[] CheckFilesInside(string folder)
        {
            if (Directory.Exists(folder))
            {
                string[] files = Directory.GetFiles(folder, "*.docx");

                foreach (string file in files)
                {
                    files.Append(file);
                }

                return files;
            }
            else
            {
                Console.WriteLine("Files does not exist.");
                return null;
            }
        }

        static string[] CheckFoldersInside(string path)
        {
            if (Directory.Exists(path))
            {
                string[] files = Directory.GetFileSystemEntries(path, "*", SearchOption.TopDirectoryOnly);

                List<string> folders = new List<string>();

                foreach (string file in files)
                {
                    folders.Add(Path.GetFullPath(file));
                }

                return folders.ToArray();
            }
            else
            {
                Console.WriteLine("Directory does not exist.");
                return null;
            }
        }
    }
}
