using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTO
{
    public class OneDayEarlyHour
    {
        //ფიზიკური მიწოდების პერიოდი
        public int Hour { get; set; }

        //2 დღით ადრე მოხმარების პროგნოზი
        public double PredictionTwoDays { get; set; }

        //2 დღით ადრე უბალანსობა
        public double UnbalanceTwoDays { get; set; }

        //ორმხრივი ხელშეკრულების ჯამი
        public double ContractSum { get; set; }

        //1 დღით ადრე უბალანსობა
        public double UnbalanceOneDay { get; set; }

        //1 დღით ადრე მოხმარების პროგნოზი
        public double PredictionOneDay { get; set; }

        //1 დღით ადრე ბაზარზე სავაჭრო ელექტროენერგიის რაოდენობა
        public double VolumeOneday { get; set; }

        //1 დღით ადრე ბაზარზე გაყიდვისთვის მინიმალური და შესყიდვისთვის მაქსიმალური ფასი
        public int Price { get; set; }

        //1 დღით ადრე ბაზარზე საპროგნოზო შემოსავალი/გასავალი
        public int IncomePrediction { get; set; }
    }
}