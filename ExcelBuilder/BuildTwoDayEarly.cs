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

        public void BuildExcel(TwoDayEarly twoDayEarly)
        {
            CreateExcelTemp(twoDayEarly.CompanyName);

            FillExcelValues(twoDayEarly);
        }

        private void FillExcelValues(TwoDayEarly twoDayEarly)
        {
            _workBook = _excel.Workbooks.Open(Path);
            _workSheet = _workBook.Worksheets[0];

            _workSheet.Cells[2][17]=twoDayEarly.CompanyName;

            List<Contractor> contractors = (List<Contractor>)twoDayEarly.Contractors;

            for(int i = 0; i<contractors.Count; i++)
            {
                FillExcelContractorValues(i, contractors[i]);
            }

            _excel.Workbooks.Close();
            _excel.Quit();
        }

        private void FillExcelContractorValues(int index, Contractor contractor)
        {
            _workSheet.Cells[2 + index][22] = 1;
            _workSheet.Cells[2 + index][23] = "A02";
            _workSheet.Cells[2 + index][24] = "8716867000016";
            _workSheet.Cells[2 + index][25] = "A03";
            //_workSheet.Cells[2 + index][26] = "???????"; trade relation
            _workSheet.Cells[2 + index][27] = "10Y1001A1001B012";
            _workSheet.Cells[2 + index][28] = "A01";
            _workSheet.Cells[2 + index][29] = "A01";
            //_workSheet.Cells[2 + index][31] = "???????"; inparty +
            _workSheet.Cells[2 + index][32] = "A01";
            //_workSheet.Cells[2 + index][33] = "????????"; outparty +
            _workSheet.Cells[2 + index][34] = "A01";
            //_workSheet.Cells[2 + index][35] = "????????????"; CapacityContractType
            //_workSheet.Cells[2 + index][36] = "???????????"; CapacityAgreementIdentification
            _workSheet.Cells[2 + index][37] = "MAW";
            //_workSheet.Cells[2 + index][38] = "???????"; TimeInterval
            _workSheet.Cells[2 + index][39] = "PT60M";

            int k = 41;
            foreach(var item in contractor.Price)
            {
                if (item > 0)
                {
                    _workSheet.Cells[2 + index][31] = contractor.BGCode; //inparty
                    _workSheet.Cells[2 + index][33] = "65YBG-TRADINGSOF"; //outparty

                    _workSheet.Cells[2 + index][k] = item;
                }
                else if(item < 0)
                {

                    _workSheet.Cells[2 + index][33] = contractor.BGCode; //outparty
                    _workSheet.Cells[2 + index][31] = "65YBG-TRADINGSOF"; //inparty

                    _workSheet.Cells[2 + index][k] = item * (-1);
                }
                else
                {
                    _workSheet.Cells[2 + index][k] = item;
                }
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
            string path = ConfigurationManager.AppSettings["SendAuctionFileAddress"] + @"\" + directoryName;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        // TODO: Add automatic temp path reader
        private string GetTempPath()
        {
            return @"C:\Users\Rabbitt\Source\Repos\S4syka\ExcelConverterProject\ExcelBuilder\Templates\Temp2.xlsm";
        }

        private string GetDestinationPath(string name)
        {
            return ConfigurationManager.AppSettings["SendTPSFileAddress"] + @"\" +
                DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]) + @"\" +
                DateTime.Now.ToString(ConfigurationManager.AppSettings["DateTimePattern"]) + "_" + name + ".xlsx";
        }
    }
}
