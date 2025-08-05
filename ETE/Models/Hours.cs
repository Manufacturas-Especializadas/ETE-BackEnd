using System;
using System.Collections.Generic;

namespace ETE.Models;

public partial class Hours
{
    public int Id { get; set; }

    public string Time { get; set; }

    public DateTime? Date { get; set; }

    public virtual ICollection<Production> Production { get; set; } = new List<Production>();

    public virtual ICollection<WorkShiftHours> WorkShiftHours { get; set; } = new List<WorkShiftHours>();
}