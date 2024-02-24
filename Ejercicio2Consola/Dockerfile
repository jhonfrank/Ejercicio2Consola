FROM mcr.microsoft.com/dotnet/aspnet:6.0 AS base
WORKDIR /app
EXPOSE 5020

ENV ASPNETCORE_URLS=http://+:5020
ENV APIbaseURL=http://api-1:5010
ENV email=franco.paredes@oechsle.pe
ENV consulta_fecha_inicio=2021-01-30
ENV consulta_fecha_final=2022-10-01

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["./*.csproj", "."]
RUN dotnet restore

COPY . .
FROM build AS publish
RUN dotnet publish -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Ejercicio2Consola.dll"]