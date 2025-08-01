using ETE.Dtos;
using ETE.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
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
                    .AsQueryable();

                if (shiftId.HasValue)
                {
                    hourIdsQuery = hourIdsQuery
                        .Where(h => _context.WorkShiftHours
                            .Any(wsh => wsh.HourId == h.Id && wsh.WorkShiftId == shiftId.Value));
                }

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
                    productionQuery = _context.Production.AsQueryable();

                    if (lineId.HasValue)
                    {
                        productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);
                    }

                    totals = await productionQuery
                        .GroupBy(p => 1)
                        .Select(g => new
                        {
                            TotalPieces = g.Sum(p => p.PieceQuantity ?? 0),
                            TotalScrap = g.Sum(p => p.Scrap ?? 0)
                        })
                        .FirstOrDefaultAsync();
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