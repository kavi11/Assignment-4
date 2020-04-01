using OpenXML.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> directories = FTP.GetDirectory(Constants.Location.FTP.BaseUrl);
            List<Student> students = new List<Student>();
            byte[] myimage = new byte[0];

            foreach (var directory in directories)
            {

                    Console.WriteLine("\n");
                    Console.WriteLine("Directory - " + directory);
                    Student student = new Student() { AbsoluteUrl = Constants.Location.FTP.BaseUrl };
                    student.FromDirectory(directory);

                    string infoFilePath = student.FullPathUrl + "/" + Constants.Location.InfoFile;
                    bool fileExists = FTP.FileExists(infoFilePath);

                    if (fileExists == true)
                    {
                        var infoBytes = FTP.DownloadFileBytes(infoFilePath);
                        string csv = Encoding.Default.GetString(infoBytes);
                        string[] csv_content = csv.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
                        if (csv_content.Length != 2)
                        {
                            Console.WriteLine("Error in CSV format");
                        }
                        else
                        {
                            student.FromCSV(csv_content[1]);
                            students.Add(student);
                        }
                    }
        
            }

            XML_Docx.CreateWordprocessingDocument(Constants.Location.DocxFile, myimage, students);
            Console.WriteLine("Docx File Created Successfully");
            XML_Excel.CreateSpreadsheetWorkbook(Constants.Location.ExcelFile, students);
            Console.WriteLine("xlsx File Created Successfully");
            UploadFile.uploadFile(Constants.Location.DocxFile, Constants.Location.FTP.WordUploadLocation);
            Console.WriteLine("Docx File Uploaded Successfully");
            UploadFile.uploadFile(Constants.Location.ExcelFile, Constants.Location.FTP.ExcelUploadLocation);
            Console.WriteLine("xlsx File Uploaded Successfully");

        }
    }
}
