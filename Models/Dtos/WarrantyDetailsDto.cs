namespace ServvistaWebAppAPI.Models.Dtos
{
    public class WarrantyDetailsDto
    {
        public string Area { get; set; }  // Changed from 'erea' to 'Area'
        public int Ns { get; set; }       // Changed to PascalCase
        public int Fs { get; set; }
        public int Ma { get; set; }
        public int Ex { get; set; }
        public int Total { get; set; }    // Added total count
    }
}