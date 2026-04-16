# Indica la imagen base que se utilizará como punto de partida
FROM python:3.11.0-slim

# Establece el directorio de trabajo dentro del contenedor
WORKDIR /app

# Copia los archivos de requerimientos al contenedor
COPY requirements.txt .

RUN python -m pip install --upgrade pip==22.3

# Instala las dependencias de la aplicación
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
RUN apt-get update && apt-get install -y locales
# Genera el locale es_ES.UTF-8
RUN echo "es_ES.UTF-8 UTF-8" > /etc/locale.gen && \
    locale-gen

# Establece las variables de entorno para el locale por defecto
ENV LANG es_ES.UTF-8
ENV LC_ALL es_ES.UTF-8

# Copia el código fuente de la aplicación al contenedor
COPY . .

# Expone el puerto en el que la aplicación escuchará (importante para Cloud Run/GKE)
EXPOSE 8001

# Define la variable de entorno para el entorno de la aplicación (opcional)
ENV APP_ENV=production

# Comando para ejecutar la aplicación cuando el contenedor se inicie
CMD ["streamlit", "run", "app_old.py", "--server.port", "8001"]