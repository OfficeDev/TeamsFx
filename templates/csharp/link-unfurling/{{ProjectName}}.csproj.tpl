<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>{{TargetFramework}}</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>

  <ItemGroup>
    <ProjectCapability Include="TeamsFx" />
  </ItemGroup>

  <ItemGroup>
    <None Remove="build/**/*" />
    <Content Remove="build/**/*" />
  </ItemGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.AspNetCore.Authentication.JwtBearer" Version="6.0.0" />
    <PackageReference Include="Microsoft.Bot.Builder" Version="4.18.1" />
    <PackageReference Include="Microsoft.Bot.Builder.Integration.AspNet.Core" Version="4.18.1" />
    <PackageReference Include="AdaptiveCards" Version="2.7.3" />
  </ItemGroup>

</Project>
