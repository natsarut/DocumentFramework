var builder = DistributedApplication.CreateBuilder(args);
builder.AddDotnetProject("webapp", "../TemplateGenerator/TemplateGenerator/TemplateGenerator.csproj");
builder.Build().Run();
