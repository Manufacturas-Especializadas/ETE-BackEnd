using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ETE.Dtos;
using ETE.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text.Json;

namespace ETE.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProductionFormController : ControllerBase
    {
        private readonly AppDbContext _context;

        public ProductionFormController(AppDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        [Route("GetListTimes")]
        public async Task<IActionResult> GetListTimes()
        {
            var list = await _context.Hours.AsNoTracking().ToListAsync();           

            return Ok(list);
        }

        [HttpGet]
        [Route("GetLines")]
        public async Task<IActionResult> GetLines()
        {
            var lines = await _context.Lines.AsNoTracking().ToListAsync();

            return Ok(lines);
        }

        [HttpGet]
        [Route("GetCodes")]
        public async Task<IActionResult> GetCodes()
        {
            var codes = await _context.Codes.AsNoTracking().ToListAsync();

            return Ok(codes);
        }

        [HttpGet]
        [Route("GetWorkShifts")]
        public async Task<IActionResult> GetWorkShifts()
        {
            var workShifts = await _context.WorkShift.AsNoTracking().ToListAsync();

            return Ok(workShifts);
        }

        [HttpGet]
        [Route("GetMachines")]
        public async Task<IActionResult> GetMachines()
        {
            var machines = await _context.Machine.AsNoTracking().ToListAsync();
            return Ok(machines);
        }

        [HttpGet]
        [Route("ValidatePartNumber/{partNumber}")]
        public async Task<IActionResult> ValidationPartNumber(string partNumber)
        {
            try
            {
                if (string.IsNullOrEmpty(partNumber))
                {
                    return BadRequest("El número de parte no puede estar vacío");
                }

                var normailzedPart = partNumber.Trim();

                var exists = await _context.PartNumberMatrix
                                .AsNoTracking()
                                .AnyAsync(p => p.PartNumber == normailzedPart);

                return Ok(new { Exists = exists });
            }
            catch(Exception ex)
            {
                return StatusCode(500, $"Error al validad el número de parte: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetProcessesByLine/{lineId}")]
        public async Task<IActionResult> GetProcessesBy(int lineId)
        {
            var processes = await _context.Process
                                .Where(p => p.LineId == lineId)
                                .AsNoTracking()
                                .ToListAsync();

            return Ok(processes);
        }

        [HttpGet]
        [Route("GetMachinesByProcess/{processId}")]
        public async Task<IActionResult> GetMachinesByProcess(int processId)
        {
            var machines = await _context.Machine
                            .Where(m => m.ProcessId == processId)
                            .AsNoTracking()
                            .ToListAsync();

            return Ok(machines);
        }

        [HttpGet]
        [Route("GetReasonByCodes/{codesId}")]
        public async Task<IActionResult> GetReasonByCodes(int codesId)
        {
            var reasons = await _context.Reason
                                .Where(r => r.CodeId == codesId)
                                .AsNoTracking()
                                .ToListAsync();

            return Ok(reasons);
        }

        [HttpGet]
        [Route("GetMachineByLine/{lineId}")]
        public async Task<IActionResult> GetMachineByLine(int lineId)
        {
            var machines = await _context.Machine
                                .Where(m => m.Process.LineId == lineId)
                                .AsNoTracking()
                                .ToListAsync();

            return Ok(machines);
        }

        [HttpGet]
        [Route("GetQualityData")]
        public async Task<IActionResult> GetQualityData(
            [FromQuery] int? lineId,
            [FromQuery] int? shiftId,
            [FromQuery] int? machineId,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                var productionQuery = _context.Production.AsQueryable();

                if (startDate.HasValue || endDate.HasValue)
                {
                    var startDateValue = startDate ?? DateTime.MinValue;
                    var endDateValue = endDate ?? DateTime.MaxValue;
                    endDateValue = endDateValue.Date.AddDays(1).AddTicks(-1);

                    productionQuery = productionQuery.Where(p =>
                        p.RegistrationDate >= startDateValue &&
                        p.RegistrationDate <= endDateValue);
                }

                if (lineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);
                }

                if (machineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);
                }

                if (shiftId.HasValue)
                {
                    productionQuery = productionQuery.Where(p =>
                        _context.WorkShiftHours.Any(wsh =>
                            wsh.HourId == p.HourId && wsh.WorkShiftId == shiftId.Value));
                }

                var totals = await productionQuery
                    .GroupBy(p => 1)
                    .Select(g => new
                    {
                        TotalPieces = g.Sum(p => p.PieceQuantity ?? 0),
                        TotalScrap = g.Sum(p => p.Scrap ?? 0)
                    })
                    .FirstOrDefaultAsync();
               
                if (totals == null)
                {
                    totals = new { TotalPieces = 0, TotalScrap = 0 };
                }

                var result = new
                {
                    GoodPieces = totals.TotalPieces - totals.TotalScrap,
                    Scrap = totals.TotalScrap,
                    TotalPieces = totals.TotalPieces,
                    QualityPercentage = totals.TotalPieces == 0 ? 0 :
                        Math.Round(((totals.TotalPieces - totals.TotalScrap) * 100.0) / totals.TotalPieces, 2)
                };

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetAvailabilityData")]
        public async Task<IActionResult> GetAvailabilityDataOptimized(
            [FromQuery] int? lineId,
            [FromQuery] int? shiftId,
            [FromQuery] int? machineId,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                IQueryable<Production> productionQuery = _context.Production.AsQueryable();

                if (startDate.HasValue || endDate.HasValue)
                {
                    var startDateValue = startDate ?? DateTime.MinValue;
                    var endDateValue = endDate ?? DateTime.MaxValue;
                    endDateValue = endDateValue.Date.AddDays(1).AddTicks(-1);

                    productionQuery = productionQuery.Where(p =>
                        p.RegistrationDate >= startDateValue &&
                        p.RegistrationDate <= endDateValue);
                }

                if (lineId.HasValue)
                    productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);

                if (machineId.HasValue)
                    productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);

                if (shiftId.HasValue)
                {
                    var validHourIds = await _context.WorkShiftHours
                        .Where(wsh => wsh.WorkShiftId == shiftId.Value)
                        .Select(wsh => wsh.HourId)
                        .ToListAsync();

                    productionQuery = productionQuery.Where(p => validHourIds.Contains(p.HourId));
                }

                var faltaDeProgramaReasonIds = await _context.Reason
                    .Where(r => EF.Functions.Like(r.Name, "%FALTA DE PROGRAMA%"))
                    .Select(r => r.Id)
                    .ToListAsync();

                var faltaDeProgramaReasonIdSet = new HashSet<int>(faltaDeProgramaReasonIds);

                var deadTimesQuery = _context.DeadTimes
                    .Where(dt => !faltaDeProgramaReasonIdSet.Contains((int)dt.ReasonId));

                var query = from p in productionQuery
                            join dt in deadTimesQuery on p.DeadTimesId equals dt.Id into dtGroup
                            from dt in dtGroup.DefaultIfEmpty()
                            select new
                            {
                                DeadTimeMinutes = dt.Minutes ?? 0
                            };

                var productionData = await query.ToListAsync();

                var totalRecords = productionData.Count;
                var totalPlannedTime = totalRecords * 60;
                var totalDeadTime = productionData.Sum(x => x.DeadTimeMinutes);

                if (totalDeadTime > totalPlannedTime)
                    totalDeadTime = totalPlannedTime;

                var totalWorkedTime = totalPlannedTime - totalDeadTime;

                double availability = totalPlannedTime > 0
                    ? (totalWorkedTime / (double)totalPlannedTime) * 100
                    : 0;

                var totalExcludedMinutes = faltaDeProgramaReasonIds
                    .Sum(id => _context.DeadTimes
                        .Where(dt => dt.ReasonId == id)
                        .Sum(dt => (int?)dt.Minutes)) ?? 0;

                var totalExcludedRecords = faltaDeProgramaReasonIds
                    .Sum(id => _context.DeadTimes
                        .Where(dt => dt.ReasonId == id)
                        .Count());

                return Ok(new
                {
                    totalTime = totalPlannedTime,
                    deadTime = totalDeadTime,
                    operatingTime = totalWorkedTime,
                    percentage = Math.Round(availability, 2),
                    debug = new
                    {
                        TotalRegistros = totalRecords,
                        TotalDeadTimeRaw = totalDeadTime,
                        TotalPlannedTime = totalPlannedTime,
                        ExcludedReasonName = "FALTA DE PROGRAMA",
                        ExcludedReasonIds = faltaDeProgramaReasonIds, 
                        TotalExcludedMinutes = totalExcludedMinutes,   
                        TotalExcludedRecords = totalExcludedRecords 
                    }
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetEfficiencyData")]
        public async Task<IActionResult> GetEfficiencyData(
            [FromQuery] int? lineId,
            [FromQuery] int? shiftId,
            [FromQuery] int? machineId,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                var productionQuery = _context.Production.AsQueryable();

                if (startDate.HasValue || endDate.HasValue)
                {
                    var startDateValue = startDate ?? DateTime.MinValue;
                    var endDateValue = endDate ?? DateTime.MaxValue;
                    endDateValue = endDateValue.Date.AddDays(1).AddTicks(-1);

                    productionQuery = productionQuery.Where(p =>
                        p.RegistrationDate >= startDateValue &&
                        p.RegistrationDate <= endDateValue);
                }

                if (lineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);
                }

                if (machineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);
                }

                if (shiftId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => _context.WorkShiftHours
                        .Any(wsh => wsh.HourId == p.HourId && wsh.WorkShiftId == shiftId.Value));
                }

                var masterQuery = _context.MasterEngineering
                    .Where(me => me.PzHr != null && me.PzHr > 0)
                    .GroupBy(me => new { me.ParentPartNumber, me.Line })
                    .Select(g => new
                    {
                        ParentPartNumber = g.Key.ParentPartNumber,
                        Line = g.Key.Line,
                        AvergaePzHr = g.Average(me => me.PzHr)
                    })
                    .AsNoTracking();

                var efficiencyQuery = productionQuery
                    .Join(masterQuery,
                        p => new { PartNumber = p.PartNumber, LineId = p.LinesId },
                        m => new { PartNumber = m.ParentPartNumber, LineId = m.Line },
                        (p, m) => new
                        {
                            p.Id,
                            p.PieceQuantity,
                            m.AvergaePzHr,
                            p.RegistrationDate,
                            p.HourId,
                            p.PartNumber,
                            p.LinesId
                        })
                    .Join(_context.Hours.AsNoTracking(),
                        x => x.HourId,
                        h => h.Id,
                        (x, h) => new
                        {
                            x.Id,
                            x.PieceQuantity,
                            x.AvergaePzHr,
                            x.RegistrationDate,
                            HourDate = h.Date,
                            x.PartNumber,
                            x.LinesId
                        });

                var groupedData = await efficiencyQuery
                    .GroupBy(x => new
                    {
                        Date = x.HourDate.HasValue ? x.HourDate.Value.Date : x.RegistrationDate.Value.Date,
                    })
                    .Select(g => new
                    {
                        Date = g.Key.Date,
                        TotalProduced = g.Sum(x => x.PieceQuantity) ?? 0,
                        TotalExpected = g.Sum(x => x.AvergaePzHr) ?? 1,
                        Efficiency = (g.Sum(x => x.PieceQuantity) ?? 0) / (double)(g.Sum(x => x.AvergaePzHr) ?? 1)
                    })
                    .OrderBy(x => x.Date)
                    .ToListAsync();

                if (!groupedData.Any() && (lineId.HasValue || machineId.HasValue || shiftId.HasValue))
                {
                    productionQuery = _context.Production.AsQueryable();

                    if (lineId.HasValue)
                        productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);

                    if (machineId.HasValue)
                        productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);

                    if (shiftId.HasValue)
                        productionQuery = productionQuery.Where(p => _context.WorkShiftHours
                            .Any(wsh => wsh.HourId == p.HourId && wsh.WorkShiftId == shiftId.Value));

                    var fallbackQuery = productionQuery
                        .Join(masterQuery,
                            p => new { PartNumber = p.PartNumber, LineId = p.LinesId },
                            m => new { PartNumber = m.ParentPartNumber, LineId = m.Line },
                            (p, m) => new
                            {
                                p.Id,
                                p.PieceQuantity,
                                m.AvergaePzHr,
                                p.RegistrationDate,
                                p.HourId
                            })
                        .Join(_context.Hours.AsNoTracking(),
                            x => x.HourId,
                            h => h.Id,
                            (x, h) => new
                            {
                                x.Id,
                                x.PieceQuantity,
                                x.AvergaePzHr,
                                x.RegistrationDate,
                                HourDate = h.Date
                            });

                    groupedData = await fallbackQuery
                        .GroupBy(x => new
                        {
                            Date = x.HourDate.HasValue ? x.HourDate.Value.Date : x.RegistrationDate.Value.Date,
                        })
                        .Select(g => new
                        {
                            Date = g.Key.Date,
                            TotalProduced = g.Sum(x => x.PieceQuantity) ?? 0,
                            TotalExpected = g.Sum(x => x.AvergaePzHr) ?? 1,
                            Efficiency = (g.Sum(x => x.PieceQuantity) ?? 0) / (double)(g.Sum(x => x.AvergaePzHr) ?? 1)
                        })
                        .OrderBy(x => x.Date)
                        .ToListAsync();
                }

                return Ok(new
                {
                    Summary = groupedData,
                    Message = "Nota: Se usa el promedio de pzHr cuando hay multiples valores para la misma combinación partNumber/Linea"
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetDeadTimeByReasonLast6Days")]
        public async Task<IActionResult> GetDeadTimeByReasonLast6Days(
            [FromQuery] int? lineId = null,
            [FromQuery] int? shiftId = null,
            [FromQuery] int? machineId = null,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                IQueryable<Production> productionQuery = _context.Production.AsQueryable();
            
                if (startDate.HasValue || endDate.HasValue)
                {
                    var startDateValue = startDate ?? DateTime.MinValue;
                    var endDateValue = endDate ?? DateTime.MaxValue;
                    endDateValue = endDateValue.Date.AddDays(1).AddTicks(-1);

                    productionQuery = productionQuery.Where(p =>
                        p.RegistrationDate >= startDateValue &&
                        p.RegistrationDate <= endDateValue);
                }

                if (lineId.HasValue)
                    productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);

                if (machineId.HasValue)
                    productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);

                if (shiftId.HasValue)
                {
                    var validHourIds = await _context.WorkShiftHours
                        .Where(wsh => wsh.WorkShiftId == shiftId.Value)
                        .Select(wsh => wsh.HourId)
                        .ToListAsync();

                    productionQuery = productionQuery.Where(p => validHourIds.Contains(p.HourId));
                }

                var relevantDeadTimeIds = await productionQuery
                    .Where(p => p.DeadTimesId != null)
                    .Select(p => p.DeadTimesId.Value)
                    .Distinct()
                    .ToListAsync();

                if (!relevantDeadTimeIds.Any())
                {
                    return Ok(new
                    {
                        labels = new string[0],
                        data = new int[0],
                        totalMinutes = 0,
                        averageMinutes = 0.0,
                        Metadata = new
                        {
                            StartDate = startDate,
                            EndDate = endDate,
                            LineFilter = lineId,
                            MachineFilter = machineId,
                            ShiftFilter = shiftId,
                            HasDateFilter = startDate.HasValue || endDate.HasValue
                        }
                    });
                }

                var result = await _context.DeadTimes
                    .Where(dt => relevantDeadTimeIds.Contains(dt.Id))
                    .GroupBy(dt => dt.Reason.Name)
                    .Select(g => new
                    {
                        Reason = g.Key ?? "Sin razón",
                        TotalMinutes = g.Sum(dt => dt.Minutes ?? 0)
                    })
                    .OrderByDescending(x => x.TotalMinutes)
                    .ToListAsync();

                var totalMinutes = result.Sum(x => x.TotalMinutes);
                var averageMinutes = totalMinutes / 6.0;

                return Ok(new
                {
                    labels = result.Select(x => x.Reason).ToArray(),
                    data = result.Select(x => x.TotalMinutes).ToArray(),
                    totalMinutes,
                    averageMinutes,
                    Metadata = new
                    {
                        StartDate = startDate,
                        EndDate = endDate,
                        LineFilter = lineId,
                        MachineFilter = machineId,
                        ShiftFilter = shiftId,
                        HasDateFilter = startDate.HasValue || endDate.HasValue
                    }
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetKeyMetrics")]
        public async Task<IActionResult> GetKeyMetrics(
            [FromQuery] int? lineId = null,
            [FromQuery] int? shiftId = null,
            [FromQuery] int? machineId = null,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {               
                DateTime effectiveStartDate = startDate ?? new DateTime(1900, 1, 1);
                DateTime effectiveEndDate = endDate.HasValue
                    ? endDate.Value.Date.AddDays(1).AddTicks(-1)
                    : DateTime.MaxValue;

                IQueryable<Production> productionQuery = _context.Production.AsQueryable();

                productionQuery = productionQuery.Where(p =>
                    p.RegistrationDate >= effectiveStartDate &&
                    p.RegistrationDate <= effectiveEndDate);

                if (lineId.HasValue)
                    productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);

                if (machineId.HasValue)
                    productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);

                if (shiftId.HasValue)
                {
                    var validHourIds = await _context.WorkShiftHours
                        .Where(wsh => wsh.WorkShiftId == shiftId.Value)
                        .Select(wsh => wsh.HourId)
                        .ToListAsync();

                    productionQuery = productionQuery.Where(p => validHourIds.Contains(p.HourId));
                }

                var faltaDeProgramaReasonIds = await _context.Reason
                    .Where(r => EF.Functions.Like(r.Name, "%FALTA DE PROGRAMA%"))
                    .Select(r => r.Id)
                    .ToListAsync();

                var relevantDeadTimeIds = await productionQuery
                    .Where(p => p.DeadTimesId != null)
                    .Select(p => p.DeadTimesId.Value)
                    .Distinct()
                    .ToListAsync();

                if (!relevantDeadTimeIds.Any())
                {
                    var totalScrap1 = await productionQuery.SumAsync(p => p.Scrap ?? 0);
                    return Ok(new
                    {
                        DeadTime = 0,
                        DeadTimeVsAvg = 0,
                        Scrap = totalScrap1,
                        ScrapVsAvg = 0,
                        Metadata = new
                        {
                            StartDate = startDate,
                            EndDate = endDate,
                            LineFilter = lineId,
                            MachineFilter = machineId,
                            ShiftFilter = shiftId,
                            HasDateFilter = startDate.HasValue || endDate.HasValue
                        }
                    });
                }

                var deadTimeQuery = _context.DeadTimes
                    .Where(dt => relevantDeadTimeIds.Contains(dt.Id) &&
                                !faltaDeProgramaReasonIds.Contains(dt.ReasonId.Value));

                var totalDeadTime = await deadTimeQuery.SumAsync(dt => dt.Minutes ?? 0);

                var totalScrap = await productionQuery.SumAsync(p => p.Scrap ?? 0);

                var avgDeadTimeQuery = _context.DeadTimes
                    .Where(dt => dt.RegistrationDate >= effectiveStartDate && 
                                dt.RegistrationDate <= effectiveEndDate &&
                                !faltaDeProgramaReasonIds.Contains(dt.ReasonId.Value));

                var avgDeadTime = await avgDeadTimeQuery.AverageAsync(dt => dt.Minutes ?? 0);

                var avgScrapQuery = _context.Production
                    .Where(p => p.RegistrationDate >= effectiveStartDate &&
                                p.RegistrationDate <= effectiveEndDate);

                var avgScrap = await avgScrapQuery.AverageAsync(p => p.Scrap ?? 0);

                return Ok(new
                {
                    DeadTime = totalDeadTime,
                    DeadTimeVsAvg = totalDeadTime - avgDeadTime,
                    Scrap = totalScrap,
                    ScrapVsAvg = totalScrap - avgScrap,
                    Metadata = new
                    {
                        StartDate = startDate,
                        EndDate = endDate,
                        LineFilter = lineId,
                        MachineFilter = machineId,
                        ShiftFilter = shiftId,
                        HasDateFilter = startDate.HasValue || endDate.HasValue
                    }
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetListProduction")]
        public async Task<IActionResult> GetLisProduction()
        {
            var list = await _context.Production
                .Select(p => new
                {
                    p.Id,
                    p.RegistrationDate,
                    Line = p.Lines.Name,
                    p.PartNumber,
                    Machine = p.Process.Machine.FirstOrDefault()!.Name,
                    Hour = p.Hour.Time,
                    p.PieceQuantity,
                    p.DeadTimes.Minutes,
                    p.Scrap
                })
                .OrderByDescending(p => p.Id)
                .AsNoTracking()
                .ToListAsync();

            return Ok(list);
        }

        [HttpPost]
        [Route("DownloadProduction")]
        public async Task<IActionResult> ExportToExcel()
        {
            var production = await _context.Production.Select(p => new
            {
                p.Id,
                p.RegistrationDate,
                p.PartNumber,
                Line = p.Lines.Name,
                Machine = p.Process.Machine.FirstOrDefault()!.Name,
                Hour = p.Hour.Time,
                p.PieceQuantity,
                p.DeadTimes.Minutes,
                p.Scrap
            })
            .AsNoTracking()
            .ToListAsync();

            using (var workBook = new XLWorkbook())
            {
                var workSheet = workBook.Worksheets.Add("Production");

                workSheet.Cell(1, 1).Value = "Fecha";
                workSheet.Cell(1, 2).Value = "Línea";
                workSheet.Cell(1, 3).Value = "Numero de parte";
                workSheet.Cell(1, 4).Value = "Máquina";
                workSheet.Cell(1, 5).Value = "Hora";
                workSheet.Cell(1, 6).Value = "Cantidad de Piezas";
                workSheet.Cell(1, 7).Value = "Tiempo Muerto (Minutos)";
                workSheet.Cell(1, 8).Value = "Scrap";

                for(int i = 0; i <  production.Count; i++)
                {
                    workSheet.Cell(i + 2, 1).Value = production[i].RegistrationDate;
                    workSheet.Cell(i + 2, 2).Value = production[i].Line;
                    workSheet.Cell(i + 2, 3).Value = production[i].PartNumber;
                    workSheet.Cell(i + 2, 4).Value = production[i].Machine;
                    workSheet.Cell(i + 2, 5).Value = production[i].Hour;
                    workSheet.Cell(i + 2, 6).Value = production[i].PieceQuantity;
                    workSheet.Cell(i + 2, 7).Value = production[i].Minutes;
                    workSheet.Cell(i + 2, 8).Value = production[i].Scrap;
                }

                var stream = new MemoryStream();
                workBook.SaveAs(stream);
                stream.Position = 0;

                var fileName = $"DatosExportados_{DateTime.Now:ddMMyyyy}.xlsx";

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        [HttpPost]
        [Route("RegisterProduction")]
        public async Task<IActionResult> RegisterProduction([FromBody] ProductionRegisterDto productionDto)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return BadRequest(ModelState);
                }

                if(productionDto.DeadTimes != null && productionDto.DeadTimes.Any())
                {
                    foreach(var dt in productionDto.DeadTimes)
                    {
                        var codeExists = await _context.Codes.AnyAsync(c => c.Id == dt.CodeId);
                        var reasonExists = await _context.Reason.AnyAsync(r => r.Id == dt.ReasonId);

                        if(!codeExists || !reasonExists)
                        {
                            return BadRequest("Código o razón de tiempo muerto no válido");
                        }
                    }
                }

                List<DeadTimes> createdDeadTimes = new List<DeadTimes>();
                if(productionDto.DeadTimes != null && productionDto.DeadTimes.Any())
                {
                    foreach(var deadTimeDto in productionDto.DeadTimes)
                    {
                        var deadTime = new DeadTimes
                        {
                            Minutes = deadTimeDto.Minutes,
                            CodeId = deadTimeDto.CodeId,
                            ReasonId = deadTimeDto.ReasonId,
                        };

                        _context.DeadTimes.Add(deadTime);
                        createdDeadTimes.Add(deadTime);
                    }
                    await _context.SaveChangesAsync();
                }

                var production = new Production
                {
                    PartNumber = productionDto.PartNumber,
                    PieceQuantity = productionDto.PieceQuantity,
                    HourId = productionDto.HourId,
                    LinesId = productionDto.LinesId,
                    ProcessId = productionDto.ProcessId,
                    MachineId = productionDto.MachineId,
                    Scrap = productionDto.Scrap,
                    ManualDate = productionDto.ManualDate,
                };

                if (createdDeadTimes.Any())
                {
                    foreach(var dt in createdDeadTimes)
                    {
                        production.DeadTimesId = dt.Id;
                    }
                }

                _context.Production.Add(production);
                await _context.SaveChangesAsync();

                return Ok(new
                {
                    Succes = true,
                    ProductionId = production.Id,
                    DeadTimesIds = createdDeadTimes.Select(dt => dt.Id).ToList()
                });
            }
            catch(Exception ex)
            {
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }
    }
}