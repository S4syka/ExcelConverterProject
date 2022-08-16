﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTO
{
    public class OneDayEarly
    {
        public IEnumerable<OneDayEarlyHour> Hour { get; set; }

        public string CompanyName { get; set; }
    }
}
