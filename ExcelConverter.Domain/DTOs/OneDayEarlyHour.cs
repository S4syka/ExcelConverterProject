using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTOs
{
    public class OneDayEarlyHour
    {
        public int Id { get; set; }
        public int OneDayEarlyExcelId { get; set; }
        public int Hour { get; set; }
        public double UnbalanceTwoDays { get; set; }
        public double PredictionTwoDays { get; set; }
        public double ContractSum { get; set; }
        public double UnbalanceOneDay { get; set; }
        public double PredictionOneDay { get; set; }
        public double VolumeOneDay { get; set; }
        public int Price { get; set; }
        public int IncomePrediction { get; set; }
    }
}
