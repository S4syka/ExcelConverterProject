using ExcelConverter.Domain.DTO;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class ReadOneDayEarly
    {
        private _Application _excel = new _Excel.Application();
        private Workbook _workBook;
        private Worksheet _workSheet;

        private double ReadCell(int i, int j)
        {
            if (_workSheet.Cells[i][j].Value != null)
            {
                return _workSheet.Cells[i][j].Value;
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
                _workBook = _excel.Workbooks.Open(item);
                _workSheet = _workBook.Worksheets[2];

                yield return new OneDayEarly() {Hour = GetDayOneDTO(),CompanyName = _workSheet.Cells[2][1].Value };

                _workBook.Close();
            }
        }
    }
}