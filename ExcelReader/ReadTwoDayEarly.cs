using ExcelConverter.Domain.DTO;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class ReadTwoDayEarly
    {
        private _Application excel = new _Excel.Application();
        private Workbook workBook;
        private Worksheet workSheet;

        private double ReadCell(int i, int j)
        {
            if (workSheet.Cells[i][j].Value != null)
            {
                return workSheet.Cells[i][j].Value;
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
                yield return Convert.ToDouble(ReadCell(index + 6, i + 7));
            }
        }

        private IEnumerable<Contractor> GetContractors()
        {
            int i = 0;
            while (true)
            {
                var temp = workSheet.Cells[i + 7][7].Value;
                if (temp == null || temp == "") break;

                yield return new Contractor()
                {
                    BGCode = Convert.ToString(ReadCell(i + 7, 4)),
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
                workBook = excel.Workbooks.Open(item);
                workSheet = workBook.Worksheets[0];

                yield return new TwoDayEarly()
                {
                    CompanyName = Convert.ToString(ReadCell(1, 1)),
                    Hour = GetDayTwoHours(),
                    Contractors = GetContractors()
                } ;

                workBook.Close();
            }
        }
    }
}
