using ETE.Models;
using System.ComponentModel.DataAnnotations;

namespace ETE.Dtos
{
    public class ProductionRegisterDto
    {
        [Required]
        public string PartNumber { get; set; }

        [Required]
        public int PieceQuantity { get; set; }
        
        public int? Scrap { get; set; }

        [Required]
        public int HourId { get; set; }

        [Required]
        public int LinesId { get; set; } 

        [Required]
        public int ProcessId { get; set; }

        [Required]
        public int MachineId { get; set; }

        public DateTime? ManualDate { get; set; }

        public List<DeadTimeDto> DeadTimes { get; set; }
    }
}