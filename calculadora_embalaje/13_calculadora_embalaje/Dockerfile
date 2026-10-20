# Contenedor de la calculadora de embalaje de KIMAK.
# Sistemas tiene que confirmar la imagen base permitida y si hay un registro
# de imágenes interno del que haya que tirar en lugar de internet.
#
# Para probarlo en local:
#   docker build -t kimak-calculadora-embalaje .
#   docker run -p 8080:80 kimak-calculadora-embalaje
#   y abrir http://localhost:8080

FROM nginx:alpine

COPY src/ /usr/share/nginx/html/

# Página de estado, para comprobar que el contenedor responde sin abrir la app.
RUN printf 'ok\n' > /usr/share/nginx/html/estado.txt

EXPOSE 80

# La aplicación no guarda datos y no necesita variables de entorno.
