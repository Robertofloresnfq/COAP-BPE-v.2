FROM python:3.10-slim

WORKDIR /app

# Actualizamos repositorios base e instalamos utilidades de compilación básicas necesarias por pandas y librerías científicas
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

COPY ./requirements.txt /app/requirements.txt

# Instalamos las dependencias de Python
RUN pip install --no-cache-dir --upgrade -r /app/requirements.txt

# Copiamos la estructura del backend y la lógica
COPY ./backend /app/backend
COPY ./logica_informes.py /app/logica_informes.py

# Exponemos el puerto 7860 (Estricto para Hugging Face Spaces)
EXPOSE 7860

# Comando para ejecutar el backend
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "7860"]