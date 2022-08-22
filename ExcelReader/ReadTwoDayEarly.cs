using ExcelConverter.Domain.DTO;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class ReadTwoDayEarly
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

        private IEnumerable<TwoDayEarlyHour> GetDayTwoHours()
        {
            for (int i = 1; i < 25; i++)
            {
                yield return new TwoDayEarlyHour()
                {
                    Hour = i,
                    Usage = Convert.ToDouble(ReadCell(2, i + 6)),
                    UnbalanceTwoDays = Convert.ToDouble(ReadCell(3, i + 6)),
                    ContractSum = Convert.ToDouble(ReadCell(4, i + 6)),
                    ContractSumBuy = Convert.ToDouble(ReadCell(5, i + 6)),
                    ContractSumSell = Convert.ToDouble(ReadCell(6, i + 6))
                };
            }
        }

        private IEnumerable<double> GetContractorPrices(int index)
        {
            for (int i = 0; i < 24; i++)
            {
                yield return Convert.ToDouble(ReadCell(index + 7, i + 7));
            }
        }

        private IEnumerable<Contractor> GetContractors()
        {
            int i = 0;
            while (true)
            {
                var temp = Convert.ToString(_workSheet.Cells[i + 7][7].Value);
                if (temp == null || temp == "") break;

                yield return new Contractor()
                {
                    BGCode = _workSheet.Cells[i + 7][4].Value,
                    Price = GetContractorPrices(i)
                };

                i++;
            }
        }

        public IEnumerable<TwoDayEarly> GetTwoDayEarlyDTOs()
        {
            string directoryPath = ConfigurationManager.AppSettings["SaveFileAddress"];
            directoryPath = directoryPath + @"\" + DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]);
            var contents = Directory.GetFiles(directoryPath, "*.xlsx");
            foreach (var item in contents)
            {
                _workBook = _excel.Workbooks.Open(item);
                _workSheet = _workBook.Worksheets[1];

                var hour = GetDayTwoHours();
                var contractors = GetContractors();

                yield return new TwoDayEarly()
                {
                    CompanyName = _workSheet.Cells[1][1].Value,
                    Hour = hour,
                    Contractors = contractors
                } ;

                _workBook.Close();
            }
        }
    }
}
