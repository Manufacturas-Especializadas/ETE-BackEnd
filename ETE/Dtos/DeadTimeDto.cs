using System.ComponentModel.DataAnnotations;

namespace ETE.Dtos
{
    public class DeadTimeDto
    {
        [Required]
        public int CodeId { get; set; }

        [Required]
        public int Minutes { get; set; }

        [Required]
        public int ReasonId { get; set; }
    }
}