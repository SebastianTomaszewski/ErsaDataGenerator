using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MetroFramework.Controls;

namespace ErsaDataGenerator
{
    public static class StringExtensions
    {
        public static string ToDateString(this DateTime s)
        {       
            return $@"{s.ToShortDateString()} 06:00:00";
        }
    }

    
}
