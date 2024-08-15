using Autofac;
using Autofac.Extensions.DependencyInjection;
using WMApp.Interfaces;
using WMApp.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddControllersWithViews().AddRazorRuntimeCompilation();

// Use Autofac to inject dependencies
// Call UseServiceProviderFactory on the Host sub property
builder.Host.UseServiceProviderFactory(new AutofacServiceProviderFactory());

// Call ConfigureContainer on the Host sub property 
builder.Host.ConfigureContainer<ContainerBuilder>(builder =>
{
    builder.RegisterType<AccessReader>().As<IAccessReader>().SingleInstance();
    builder.RegisterType<ExcelReader>().As<IExcelReader>().SingleInstance();
    builder.RegisterType<FormFiller>().As<IFormFiller>().SingleInstance();
});


License.LicenseKey = "IRONSUITE.XNFKMK3.FEMAILTOR.COM.13749-2033057F36-ANULSNSMO4NYH3SQ-CV5TGK5ESHU7-RDSQUUZ5PO3A-RNB4JFU7COPO-ORKUERR6IFAE-O4S2IXA4D33U-LCM4ZT-TGNT34Y453CNEA-DEPLOYMENT.TRIAL-BQFX75.TRIAL.EXPIRES.14.SEP.2024";
var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
