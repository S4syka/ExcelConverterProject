using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseHelper;
using ExcelConverter.Domain.DTOs;

namespace ExcelConverter.Repository
{
    public class OneDayEarlyExcelsRepository : BaseRepository<OneDayEarlyExcel>
    {
        public OneDayEarlyExcelsRepository(Database database) : base(database)
        {

        }
    }
}
