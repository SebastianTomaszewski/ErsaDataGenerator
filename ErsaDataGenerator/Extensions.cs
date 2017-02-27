using System;

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
