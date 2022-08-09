using System.IO;
using System.Configuration;
using ExcelConverter.Domain.DTO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class ReadDayOne
    {
        _Application excel = new _Excel.Application();
        Workbook workBook;
        Worksheet workSheet;

        private string Path { get; set; }

        public ReadDayOne(string path)
        {
            this.Path = path;
            workBook = excel.Workbooks.Open(path);
            workSheet = workBook.Worksheets[2];
        }

        private double ReadCell(int i, int j)
        {
            if (workSheet.Cells[i][j].Value != null)
            {
                return workSheet.Cells[i][j].Value;
            }
            else return 0;
        }

        private IEnumerable<OneDayEarly> GetDayOneDTO()
        {
            for (int i = 1; i < 25; i++)
            {
                yield return new OneDayEarly()
                {
                    Hour = i,
                    PredictionTwoDays = ReadCell(2, i + 6),
                    UnbalanceTwoDays = ReadCell(3, i + 6),
                    ContractSum = ReadCell(4, i + 6),
                    UnbalanceOneDay = ReadCell(5, i + 6),
                    PredictionOneDay = ReadCell(6, i + 6),
                    VolumeOneday = ReadCell(7, i + 6),
                    Price = Convert.ToInt32(ReadCell(8, i + 6)),
                    IncomePrediction = Convert.ToInt32(ReadCell(9, i + 6))
                };
            }
        }

        public IEnumerable<IEnumerable<OneDayEarly>> GetDayOneDTOs()
        {
            string directoryPath = ConfigurationManager.AppSettings["SaveFileAddress"];
            directoryPath = directoryPath + @"\" + DateTime.Now.ToString("dd-MM-yyyy");
            var contents = Directory.GetFiles(directoryPath,"*.xml");
            foreach (var item in contents)
            {
                workBook = excel.Workbooks.Open(item);
                workSheet = workBook.Worksheets[2];

                yield return GetDayOneDTO();

                workBook.Close();
            }
        }
    }
}