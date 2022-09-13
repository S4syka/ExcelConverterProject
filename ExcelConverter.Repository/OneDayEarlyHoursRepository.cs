using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseHelper;
using ExcelConverter.Domain.DTOs;

namespace ExcelConverter.Repository
{
    public class OneDayEarlyHoursRepository : BaseRepository<OneDayEarlyHour>
    {
        public OneDayEarlyHoursRepository(Database database) : base(database)
        {

        }
    }
}
