using ExcelConverter.Domain.DTO;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class ReadOneDayEarly
    {
        _Application excel = new _Excel.Application();
        Workbook workBook;
        Worksheet workSheet;

        private double ReadCell(int i, int j)
        {
            if (workSheet.Cells[i][j].Value != null)
            {
                return workSheet.Cells[i][j].Value;
            }
            else return 0;
        }

        private IEnumerable<OneDayEarlyHour> GetDayOneDTO()
        {
            for (int i = 1; i < 25; i++)
            {
                yield return new OneDayEarlyHour()
                {
                    Hour = i,
                    PredictionTwoDays = Convert.ToDouble(ReadCell(2, i + 6)),
                    UnbalanceTwoDays = Convert.ToDouble(ReadCell(3, i + 6)),
                    ContractSum = Convert.ToDouble(ReadCell(4, i + 6)),
                    UnbalanceOneDay = Convert.ToDouble(ReadCell(5, i + 6)),
                    PredictionOneDay = Convert.ToDouble(ReadCell(6, i + 6)),
                    VolumeOneday = Convert.ToDouble(ReadCell(7, i + 6)),
                    Price = Convert.ToInt32(ReadCell(8, i + 6)),
                    IncomePrediction = Convert.ToInt32(ReadCell(9, i + 6))
                };
            }
        }

        public IEnumerable<OneDayEarly> GetDayOneDTOs()
        {
            string directoryPath = ConfigurationManager.AppSettings["SaveFileAddress"];
            directoryPath = directoryPath + @"\" + DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]);
            var contents = Directory.GetFiles(directoryPath, "*.xlsx");
            foreach (var item in contents)
            {
                workBook = excel.Workbooks.Open(item);
                workSheet = workBook.Worksheets[2];

                yield return new OneDayEarly(GetDayOneDTO(), workSheet.Cells[2][1].Value);

                workBook.Close();
            }
        }
    }
}