using ETE.Models;
using ETE.Services;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
using System.Net;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

//Services
builder.Services.AddSingleton<EmailService>();

var connection = builder.Configuration.GetConnectionString("Connection");
builder.Services.AddDbContext<AppDbContext>(options =>
{
    options.UseSqlServer(connection);
});

var allowedConnection = builder.Configuration.GetValue<string>("OrigenesPermitidos")!.Split(',');

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowSpecificOrigins", policy =>
    {
        policy.WithOrigins(allowedConnection)
                        .AllowAnyHeader()
                        .AllowAnyMethod();
    });
});

//var ipAllowedStr = builder.Configuration.GetValue<string>("IpPermitida");
//if(!IPAddress.TryParse(ipAllowedStr, out var iPAddress))
//{
//    throw new Exception("La ipPermitida en appsettings.json no es valida");
//}

var app = builder.Build();

//app.Use(async (context, next) =>
//{
//    var remoteIp = context.Connection.RemoteIpAddress;

//    if(context.Request.Headers.TryGetValue("X-Forwarded-For", out var forwardedHeader))
//    {
//        var firstIp = forwardedHeader.ToString().Split(',').FirstOrDefault();
//        if(IPAddress.TryParse(firstIp, out var parsedIp))
//        {
//            remoteIp = parsedIp;
//        }
//    }

//    Debug.WriteLine($"Ip conectada: {remoteIp}");

//    if (!remoteIp.Equals(iPAddress))
//    {
//        context.Response.StatusCode = 403;
//        await context.Response.WriteAsync("Acceso denegado: IP no autorizada");

//        return;
//    }

//    await next();
//});

app.UseSwagger();
app.UseSwaggerUI();

app.UseCors("AllowSpecificOrigins");

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
