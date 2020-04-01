using System;
using System.Collections.Generic;
using System.Text;

namespace OpenXML.Models
{
    class Constants
    {
        public readonly Student Student = new Student { StudentId = "200447599", FirstName = "KavirajSingh", LastName = "Jon" };
        public class Location
        {
            public readonly static string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            public readonly static string ExePath = Environment.CurrentDirectory;

            public readonly static string ContentFolder = $"{ExePath}\\..\\..\\..\\Content";
            public readonly static string ImagesFolder = $"{ContentFolder}\\Image";
            public readonly static string DataFolder = $"{ContentFolder}";
            public readonly static string DataFolderImage = $"{ContentFolder}\\Image\\myimage.jpg"; 

            public const string InfoFile = "info.csv";
            public const string ImageFile = "myimage.jpg";

            public readonly static string ExcelFile = $"{ContentFolder}\\info.xlsx";
            public readonly static string DocxFile = $"{ContentFolder}\\info.docx";

            public class FTP
            {
                public const string Username = @"bdat100119f\bdat1001";
                public const string Password = "bdat1001";

                public const string BaseUrl = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914";
                public const string MyDirectory = "/200447599 KavirajSingh Jon";
                public const string ImageUrl = BaseUrl + MyDirectory + "/myimage.jpg";

                public const string ExcelUploadLocation = BaseUrl + MyDirectory + "/info.xlsx";
                public const string WordUploadLocation = BaseUrl + MyDirectory + "/info.docx";


                public const int OperationPauseTime = 10000;
            }
        }
    }
}
