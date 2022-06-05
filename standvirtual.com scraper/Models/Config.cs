using System.Collections.Generic;

namespace standvirtual.com_scraper.Models
{
    public class Config
    {
        public List<Make> Makes { get; set; } = new List<Make>();
        public List<string> Prices { get; set; }
        public List<string> Dates { get; set; }
        public List<string> MileAges { get; set; }
        public List<string> BatteriesPower { get; set; }
    }
}