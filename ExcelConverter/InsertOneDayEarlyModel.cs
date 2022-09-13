using ExcelConverter.Domain.DTOs;
using ExcelConverter.Domain.Model;
using ExcelConverter.Repository;
using System.Collections.ObjectModel;

namespace ExcelConverter.Domain
{
    public class InsertOneDayEarlyModel
    {
        private UnitOfWork _unitOfWork;

        public OneDayEarlyExcel OneDayEarlyExcel { get; private set; }
        public ICollection<OneDayEarlyHour> OneDayEarlyHours { get; private set; }

        public InsertOneDayEarlyModel(OneDayEarlyModel model)
        {
            OneDayEarlyExcel = new();
            OneDayEarlyHours = new Collection<OneDayEarlyHour>();

            _unitOfWork = new UnitOfWork();

            OneDayEarlyExcel.CompanyName = model.CompanyName;
            OneDayEarlyExcel.ExcelDate = DateTime.Now;

            int id = _unitOfWork.OneDayEarlyExcelsRepository.Insert(OneDayEarlyExcel);
            OneDayEarlyExcel.Id = id;

            foreach (var item in model.Hour)
            {
                OneDayEarlyHours.Add(new OneDayEarlyHour
                {
                    OneDayEarlyExcelId = id,
                    Hour = item.Hour,
                    PredictionTwoDays = item.PredictionTwoDays,
                    UnbalanceTwoDays = item.UnbalanceTwoDays,
                    ContractSum = item.ContractSum,
                    UnbalanceOneDay = item.UnbalanceOneDay,
                    PredictionOneDay = item.PredictionOneDay,
                    VolumeOneDay = item.VolumeOneday,
                    Price = item.Price,
                    IncomePrediction = item.IncomePrediction
                });
            }

            foreach (var item in OneDayEarlyHours)
            {
                _unitOfWork.OneDayEarlyHoursRepository.Insert(item);
            }
        }
    }
}
