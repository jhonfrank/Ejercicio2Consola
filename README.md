# Ejercicio2Consola

Para ejecutar en docker:

Port: 5020

ENV:--
      - ASPNETCORE_URLS=http://+:5020--
      - APIbaseURL=http://localhost:5010--
      - email=jhonfrank.ae@outlook.com--
      - consulta_fecha_inicio=2021-01-30 **format yyy-mm-dd**--
      - consulta_fecha_final=2022-10-01 **format yyy-mm-dd**--

## Docker exec

docker build -t consola--
docker run --network=nt0 -p 5020:5020 --name consola-1 consola