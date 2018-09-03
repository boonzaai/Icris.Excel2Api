FROM microsoft/dotnet:2.1-aspnetcore-runtime AS base
WORKDIR /app
EXPOSE 60090
EXPOSE 44344

FROM microsoft/dotnet:2.1-sdk AS build
WORKDIR /src
COPY Icris.Excel2Api.CoreWeb/Icris.Excel2Api.CoreWeb.csproj Icris.Excel2Api.CoreWeb/
RUN dotnet restore Icris.Excel2Api.CoreWeb/Icris.Excel2Api.CoreWeb.csproj
COPY . .
WORKDIR /src/Icris.Excel2Api.CoreWeb
RUN dotnet build Icris.Excel2Api.CoreWeb.csproj -c Release -o /app

FROM build AS publish
RUN dotnet publish Icris.Excel2Api.CoreWeb.csproj -c Release -o /app

FROM base AS final
WORKDIR /app
COPY --from=publish /app .
ENTRYPOINT ["dotnet", "Icris.Excel2Api.CoreWeb.dll"]
