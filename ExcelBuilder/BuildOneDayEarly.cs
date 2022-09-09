using ExcelConverter.Domain.Model;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelBuilder
{
    public class BuildOneDayEarly
    {
        private _Application _excel = new _Excel.Application();
        private Workbook _workBook;
        private Worksheet _workSheet;

        private string Path { get; set; }
        private string Source { get; set; }

        public void BuildExcel(OneDayEarlyModel oneDayEarly)
        {
            CreateExcelTemp(oneDayEarly.CompanyName);

            List<double> firstRow = GetFirstRow(oneDayEarly.Hour);

            FillExcelFile(firstRow, oneDayEarly.Hour, oneDayEarly.CompanyName);
        }

        private void CreateExcelTemp(string name)
        {
            Source = GetTempPath();
            Path = GetDestinationPath(name);
            CreateNewDirectory();

            if (!File.Exists(Path))
            {
                File.Copy(Source, Path);
            }
        }

        private string CreateNewDirectory()
        {
            string directoryName = DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]);
            string path = ConfigurationManager.AppSettings["SendAuctionFileAddress"] + @"\" + directoryName;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        private string GetTempPath()
        {
            return ConfigurationManager.AppSettings["TempForAuctionPath"];
        }

        private string GetDestinationPath(string name)
        {
            return ConfigurationManager.AppSettings["SendAuctionFileAddress"] + @"\" +
                DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]) + @"\" +
                DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]) + "_" + name + ".xlsx";
        }

        private List<double> GetFirstRow(IEnumerable<OneDayEarlyHourModel> data)
        {
            List<double> rawRow = new List<double>();
            List<double> row = new List<double>();

            rawRow.Add(-200.00);
            rawRow.Add(1000.00);

            foreach (var item in data)
            {
                if (item.VolumeOneday > 0)
                {
                    rawRow.Add(item.Price + 0.00);
                    rawRow.Add(item.Price + 0.01);
                }
                else if (item.VolumeOneday < 0)
                {
                    rawRow.Add(item.Price + 0.00);
                    rawRow.Add(item.Price - 0.01);
                }
            }

            rawRow.Sort();

            row.Add(-200.00);
            for (int i = 1; i < rawRow.Count(); i++)
            {
                if (rawRow[i] != rawRow[i - 1]) row.Add(rawRow[i]);
            }

            return row;
        }

        private void FillExcelFile(List<double> firstRow, IEnumerable<OneDayEarlyHourModel> data, string companyName)
        {
            _workBook = _excel.Workbooks.Open(Path);
            _workSheet = _workBook.Worksheets[1];

            for (int k = 2; k < firstRow.Count + 2; k++)
            {
                _workSheet.Cells[k][8].Value = firstRow[k - 2];
            }

            int i = 9;
            foreach (var item in data)
            {
                for (int j = 2; j < firstRow.Count + 2; j++)
                {
                    if (item.VolumeOneday > 0 && item.Price >= firstRow[j - 2])
                    {
                        SaveCell(i, j, item.VolumeOneday.ToString());
                    }
                    else
                    if (item.VolumeOneday < 0 && item.Price <= firstRow[j - 2])
                    {
                        SaveCell(i, j, item.VolumeOneday.ToString());
                    }
                    else SaveCell(i, j, "0");
                }
                i++;
            }

            _workSheet.Name=companyName + "  GE Curve";

            SaveCell(5, 2, companyName + " ");

            _excel.Workbooks.Close();
            _excel.Quit();
        }

        private void SaveCell(int j, int i, string value)
        {
            _workSheet.Cells[i][j].Value = value;
        }
    }
}