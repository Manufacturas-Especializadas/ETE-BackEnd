using ETE.Dtos;
using ETE.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
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
                startDate ??= DateTime.Now.AddDays(-30);
                endDate ??= DateTime.Now;

                endDate = endDate.Value.Date.AddDays(1).AddTicks(-1);

                var hourIdsQuery = _context.Hours
                    .Where(h => h.Date >= startDate && h.Date <= endDate)
                    .AsNoTracking()
                    .AsQueryable();

                //if (shiftId.HasValue)
                //{
                //    hourIdsQuery = hourIdsQuery
                //        .Where(h => _context.WorkShiftHours
                //            .Any(wsh => wsh.HourId == h.Id && wsh.WorkShiftId == shiftId.Value));
                //}

                var hourIds = await hourIdsQuery
                    .Select(h => h.Id)
                    .ToListAsync();

                Console.WriteLine($"IDs de horas encontrados: {hourIds.Count}");

                var productionQuery = _context.Production
                    .Where(p => hourIds.Contains(p.HourId))
                    .AsQueryable();

                if (lineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);
                }

                if (machineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);
                }

                var totals = await productionQuery
                    .GroupBy(p => 1)
                    .Select(g => new
                    {
                        TotalPieces = g.Sum(p => p.PieceQuantity ?? 0),
                        TotalScrap = g.Sum(p => p.Scrap ?? 0)
                    })
                    .FirstOrDefaultAsync();

                if (totals == null || (totals.TotalPieces == 0 && totals.TotalScrap == 0))
                {                    
                    if (!lineId.HasValue && !machineId.HasValue && !shiftId.HasValue)
                    {
                        productionQuery = _context.Production.AsQueryable();
                        totals = await productionQuery
                            .GroupBy(p => 1)
                            .Select(g => new
                            {
                                TotalPieces = g.Sum(p => p.PieceQuantity ?? 0),
                                TotalScrap = g.Sum(p => p.Scrap ?? 0)
                            })
                            .FirstOrDefaultAsync();
                    }
                }

                var result = new
                {
                    GoodPieces = totals?.TotalPieces - totals?.TotalScrap ?? 0,
                    Scrap = totals?.TotalScrap ?? 0,
                    TotalPieces = totals?.TotalPieces ?? 0,
                    QualityPercentage = totals == null || totals.TotalPieces == 0 ? 0 :
                        Math.Round(((totals.TotalPieces - totals.TotalScrap) * 100.0) / totals.TotalPieces, 2)
                };

                Console.WriteLine($"Resultados: {JsonSerializer.Serialize(result)}");

                return Ok(result);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetEfficiencyData")]
        public async Task<IActionResult> GetEfficiencyData(
            [FromQuery] int? lineId,
            [FromQuery] int? shiftId,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                endDate ??= DateTime.Now;
                endDate = endDate.Value.Date.AddDays(1).AddTicks(-1);

                var hourCount = await _context.Hours
                    .Where(h => h.Date >= startDate && h.Date <= endDate)
                    .CountAsync();

                Console.WriteLine($"Total de horas en rango: {hourCount}");

                var query = from p in _context.Production
                            join h in _context.Hours on p.HourId equals h.Id
                            join me in _context.MasterEngineering on
                                new { PartNumber = p.PartNumber, LineId = p.LinesId }
                                equals
                                new { PartNumber = me.ChildPartNumber, LineId = me.Line }
                            into meJoin
                            from me in meJoin.DefaultIfEmpty()
                            where h.Date >= startDate && h.Date <= endDate
                            select new
                            {
                                p.PartNumber,
                                p.LinesId,
                                p.PieceQuantity,
                                Expected = me != null ? me.PzHr : null
                            };

                if (lineId.HasValue)
                {
                    query = query.Where(x => x.LinesId == lineId.Value);
                }

                //if (shiftId.HasValue)
                //{
                //    query = query.Where(x => _context.WorkShiftHours
                //        .Any(wsh => wsh.HourId == x.HourId && wsh.WorkShiftId == shiftId.Value));
                //}

                var results = await query.ToListAsync();

                Console.WriteLine($"Registros encontrados: {results.Count}");
                Console.WriteLine($"Ejemplo de primeros registros: {JsonSerializer.Serialize(results.Take(3))}");

                var validResults = results
                    .Where(x => x.Expected.HasValue && x.Expected > 0)
                    .ToList();

                var totalProduced = validResults.Sum(x => x.PieceQuantity ?? 0);
                var totalExpected = validResults.Sum(x => x.Expected ?? 0);
                var efficiency = totalExpected > 0 ? Math.Round((totalProduced * 100.0) / totalExpected, 2) : 0;

                var missingData = results
                    .Where(x => !x.Expected.HasValue || x.Expected <= 0)
                    .GroupBy(x => new { x.PartNumber, x.LinesId })
                    .Select(g => new
                    {
                        g.Key.PartNumber,
                        g.Key.LinesId,
                        Count = g.Count()
                    })
                    .ToList();

                Console.WriteLine($"Registros sin target: {missingData.Count}");
                if (missingData.Any())
                {
                    Console.WriteLine($"Top partNumbers sin target: {JsonSerializer.Serialize(missingData.Take(5))}");
                }

                return Ok(new
                {
                    TotalProduced = totalProduced,
                    TotalExpected = totalExpected,
                    Efficiency = efficiency,
                    HourCount = hourCount,
                    Details = validResults.Take(10),
                    MissingDataCount = missingData.Count,
                    DebugInfo = new
                    {
                        TotalRecords = results.Count,
                        RecordsWithTarget = validResults.Count,
                        DateRange = $"{startDate} - {endDate}"
                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error completo: {ex.ToString()}");
                return StatusCode(500, new
                {
                    Error = ex.Message,
                    StackTrace = ex.StackTrace,
                    InnerException = ex.InnerException?.Message
                });
            }
        }

        [HttpGet]
        [Route("GetAvailabilityData")]
        public async Task<IActionResult> GetAvailabilityDataNoDateFilter(
            [FromQuery] int? lineId,
            [FromQuery] int? shiftId,
            [FromQuery] int? machineId)
        {
            try
            {
                var hourQuery = _context.Hours.AsQueryable();

                if (shiftId.HasValue)
                {
                    hourQuery = hourQuery
                        .Where(h => _context.WorkShiftHours
                            .Any(wsh => wsh.HourId == h.Id && wsh.WorkShiftId == shiftId.Value));
                }

                var filteredHourIds = await hourQuery.Select(h => h.Id).ToListAsync();

                if (!filteredHourIds.Any())
                {
                    return Ok(new
                    {
                        totalTime = 0,
                        deadTime = 0,
                        availableTime = 0,
                        percentage = 0
                    });
                }

                int totalMinutes = filteredHourIds.Count * 60;

                var deadTimesQuery = _context.Production
                    .Where(p => filteredHourIds.Contains(p.HourId))
                    .Join(_context.DeadTimes,
                        p => p.DeadTimesId,
                        dt => dt.Id,
                        (p, dt) => new { dt.Minutes, p.LinesId, p.MachineId });

                if (lineId.HasValue)
                {
                    deadTimesQuery = deadTimesQuery.Where(x => x.LinesId == lineId.Value);
                }

                if (machineId.HasValue)
                {
                    deadTimesQuery = deadTimesQuery.Where(x => x.MachineId == machineId.Value);
                }

                int totalDeadTime = await deadTimesQuery.SumAsync(x => x.Minutes ?? 0);
                int availableTime = totalMinutes - totalDeadTime;
                double availabilityPercentage = totalMinutes > 0
                    ? Math.Round((double)availableTime / totalMinutes * 100, 2)
                    : 0;

                return Ok(new
                {
                    totalTime = totalMinutes,
                    deadTime = totalDeadTime,
                    availableTime = availableTime,
                    percentage = availabilityPercentage
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al calcular la disponibilidad: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetDeadTimeByReasonLast6Days")]
        public async Task<IActionResult> GetDeadTimeByReasonLast6Days([FromQuery] int? lineId = null)
        {
            try
            {
                var endDate = DateTime.Today;
                var startDate = endDate.AddDays(-6);

                var result = await _context.Production
                    .Where(p => p.Hour.Date >= startDate &&
                               p.Hour.Date <= endDate &&
                               p.DeadTimesId != null &&
                               (lineId == null || p.LinesId == lineId))
                    .GroupBy(p => p.DeadTimes.Reason.Name)
                    .Select(g => new {
                        Reason = g.Key ?? "Sin razón",
                        TotalMinutes = g.Sum(p => p.DeadTimes.Minutes ?? 0)
                    })
                    .OrderBy(x => x.TotalMinutes)
                    .ToListAsync();
               
                if (!result.Any())
                {
                    result = await _context.DeadTimes
                        .GroupBy(dt => dt.Reason.Name)
                        .Select(g => new {
                            Reason = g.Key ?? "Sin razón",
                            TotalMinutes = g.Sum(dt => dt.Minutes ?? 0)
                        })
                        .OrderBy(x => x.TotalMinutes)
                        .ToListAsync();
                }

                var totalMinutes = result.Sum(x => x.TotalMinutes);
                var averageMinutes = totalMinutes / 6.0;

                return Ok(new
                {
                    labels = result.Select(x => x.Reason).ToArray(),
                    data = result.Select(x => x.TotalMinutes).ToArray(),
                    totalMinutes,
                    averageMinutes
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
            [FromQuery] DateTime? startDate = null, 
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                var endDateValue = endDate ?? DateTime.Now;
                var startDateValue = startDate ?? endDateValue.AddDays(-7);

                var totalDeadTime = await _context.DeadTimes
                    .SumAsync(dt => dt.Minutes ?? 0);

                var avgDeadTime = await _context.DeadTimes
                    .AverageAsync(dt => dt.Minutes ?? 0);

                var totalScrap = await _context.Production
                     .Where(p => p.Scrap != null)
                     .SumAsync(p => p.Scrap ?? 0);

                var avgScrap = await _context.Production
                    .Where(p => p.Scrap != null)
                    .AverageAsync(p => p.Scrap ?? 0);

                return Ok(new
                {
                    DeadTime = totalDeadTime,
                    DeadTimeVsAvg = totalDeadTime - avgDeadTime,
                    Scrap = totalScrap,
                    ScrapVsAvg = totalScrap - avgScrap,
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
                    Machine = p.Process.Machine.FirstOrDefault()!.Name,
                    Hour = p.Hour.Time,
                    p.PieceQuantity,
                    p.DeadTimes.Minutes,
                    p.Scrap
                })
                .AsNoTracking()
                .ToListAsync();

            return Ok(list);
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