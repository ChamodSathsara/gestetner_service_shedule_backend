using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using ServvistaWebAppAPI.Classes;
using ServvistaWebAppAPI.Repositories;
using ServvistaWebAppAPI.Services;
using System.Security.Claims;
using System.Text;

var builder = WebApplication.CreateBuilder(args);

// --------------------
// CORS (React + SignalR)
// --------------------
builder.Services.AddCors(options =>
{
    //options.AddPolicy("AllowOrigin", policy =>
    //{
    //    policy
    //        .SetIsOriginAllowed(_ => true)          // ⚠️ Only for testing! Works with credentials
    //        .AllowAnyHeader()
    //        .AllowAnyMethod()
    //        .AllowCredentials();
    //});

    options.AddPolicy("AllowOrigin", policy =>
    {
        policy.WithOrigins(
            "http://localhost:5173",
            "http://localhost:3000",
            "https://gestetner-service-schedule-4cse.vercel.app",
            "https://gestetner-service-schedule.vercel.app",
            "https://gestetner-service-schedule-git-227f58-chamodsathsaras-projects.vercel.app",
            "https://servvistagcp-001-site15.anytempurl.com"

        // Add more preview branches if needed, e.g. "https://*.vercel.app"
        )
        .AllowAnyHeader()
        .AllowAnyMethod()
        .AllowCredentials();
    });


    // Recommended for production (uncomment when ready + add your real domains)

});

// --------------------
// (No background polling - AddCall triggers /api/internal/notify/job-assigned
// immediately when a job is created. See Controllers/InternalNotifyController.cs)
// --------------------

// --------------------
// Controllers & Services
// --------------------
builder.Services.AddControllers();
builder.Services.AddScoped<UserRepository>();
builder.Services.AddScoped<JwtTokenService>();
builder.Services.AddScoped<ITechnicianPerformanceService, TechnicianPerformanceService>();
builder.Services.AddScoped<IBreakdownServices, BreakdownServices>();
builder.Services.AddScoped<IServiceSchedule, ServiceScheduleService>();
builder.Services.AddHttpContextAccessor();
builder.Services.AddScoped<ITenantService, TenantService>();
builder.Services.AddScoped<IManagementRepository, ManagementServices>();
builder.Services.AddScoped<IReportRepository, ReportService>();
builder.Services.AddScoped<PushSubscriptionRepository>();
builder.Services.AddScoped<JobNotificationService>();

// --------------------
// Authentication (JWT)
// --------------------
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,
            ValidIssuer = builder.Configuration["Jwt:Issuer"],
            ValidAudience = builder.Configuration["Jwt:Audience"],
            IssuerSigningKey = new SymmetricSecurityKey(
                Encoding.UTF8.GetBytes(builder.Configuration["Jwt:Key"] ?? throw new InvalidOperationException("Jwt:Key is missing"))
            ),
            NameClaimType = ClaimTypes.Name
        };
    });

builder.Services.AddAuthorization();

// --------------------
// Swagger
// --------------------
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// --------------------
// Middleware pipeline — CORRECT ORDER
// --------------------
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

// Optional: Remove or keep — Railway usually handles HTTPS, so this can sometimes cause warnings
// app.UseHttpsRedirection();

app.UseRouting();                    // ← Must come BEFORE UseCors

// 🔥 FIXED: CORS now in the correct position (after Routing, before Auth + Endpoints)
app.UseCors("AllowOrigin");

app.UseAuthentication();             // Authentication before Authorization
app.UseAuthorization();

// --------------------
// Endpoints — these must come AFTER UseCors
// --------------------
app.MapControllers();

app.Run();