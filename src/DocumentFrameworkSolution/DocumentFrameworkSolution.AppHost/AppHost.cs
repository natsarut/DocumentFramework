var builder = DistributedApplication.CreateBuilder(args);
builder.AddDotnetProject("webapp", "../TemplateGenerator/TemplateGenerator/TemplateGenerator.csproj").WithHttpHealthCheck("/health");
builder.Build().Run();
