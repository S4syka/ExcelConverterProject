using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using ExcelConverter.Domain.DTO;
using System.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBuilder
{
    public   class BuildTwoDayEarly
    {
        private _Application _excel = new _Excel.Application();
        private Workbook _workBook;
        private Worksheet _workSheet;

        private string Path { get; set; }
        private string Source { get; set; }

        public void BuildExcel(TwoDayEarlyModel twoDayEarly)
        {
            CreateExcelTemp(twoDayEarly.CompanyName);

            FillExcelValues(twoDayEarly);
        }

        private void FillExcelValues(TwoDayEarlyModel twoDayEarly)
        {
            _workBook = _excel.Workbooks.Open(Path);
            _workSheet = _workBook.Worksheets[1];

            _workSheet.Cells[2][17]=twoDayEarly.CompanyName;

            int i = 0;
            foreach(var item in twoDayEarly.Contractors)
            {
                FillExcelContractorValues(i, item, twoDayEarly.CompanyName);
                i++;
            }

            _excel.Workbooks.Close();
            _excel.Quit();
        }

        private void FillExcelContractorValues(int index, ContractorModel contractor, string portfolio)
        {
            _workSheet.Cells[2 + index][21] = "TS00" +$"{index+1}";
            _workSheet.Cells[2 + index][22] = 1;
            _workSheet.Cells[2 + index][23] = "A02";
            _workSheet.Cells[2 + index][24] = "'8716867000016";
            _workSheet.Cells[2 + index][25] = "A03";
            //_workSheet.Cells[2 + index][26] = "???????"; trade relation
            _workSheet.Cells[2 + index][27] = "10Y1001A1001B012";
            _workSheet.Cells[2 + index][28] = "A01";
            _workSheet.Cells[2 + index][29] = "10Y1001A1001B012";
            _workSheet.Cells[2 + index][30] = "A01";
            //_workSheet.Cells[2 + index][31] = "???????"; inparty +
            _workSheet.Cells[2 + index][32] = "A01";
            //_workSheet.Cells[2 + index][33] = "????????"; outparty +
            _workSheet.Cells[2 + index][34] = "A01";
            //_workSheet.Cells[2 + index][35] = "????????????"; CapacityContractType
            //_workSheet.Cells[2 + index][36] = "???????????"; CapacityAgreementIdentification
            _workSheet.Cells[2 + index][37] = "MAW";
            _workSheet.Cells[2 + index][38] = "2022-08-08T22:00Z/2022-08-09T22:00Z";
            _workSheet.Cells[2 + index][39] = "PT60M";

            int k = 41;
            foreach(var item in contractor.Price)
            {
                if (item > 0)
                {
                    _workSheet.Cells[2 + index][33] = contractor.BGCode; //inparty
                    _workSheet.Cells[2 + index][31] = portfolio; //outparty

                    _workSheet.Cells[2 + index][k] = item;
                }
                else if(item < 0)
                {
                    _workSheet.Cells[2 + index][31] = contractor.BGCode; //outparty
                    _workSheet.Cells[2 + index][33] = portfolio; //inparty

                    _workSheet.Cells[2 + index][k] = item * (-1);
                }
                else
                {
                    _workSheet.Cells[2 + index][k] = item;
                }
                k++;
            }
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
            string path = ConfigurationManager.AppSettings["SendTPSFileAddress"] + @"\" + directoryName;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        private string GetTempPath()
        {
            return ConfigurationManager.AppSettings["TempForTPSPath"];
        }

        private string GetDestinationPath(string name)
        {
            return ConfigurationManager.AppSettings["SendTPSFileAddress"] + @"\" +
                DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]) + @"\" +
                DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]) + "_" + name + ".xlsm";
        }
    }
}
