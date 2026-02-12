# Usamos una imagen oficial de Python que ya tiene lo básico
FROM mcr.microsoft.com/playwright/python:v1.40.0-jammy

# Directorio de trabajo dentro del contenedor
WORKDIR /app

# Copiamos los archivos de tu proyecto al contenedor
COPY requirements.txt .
COPY . .

# Instalamos las librerías de Python
RUN pip install --no-cache-dir -r requirements.txt

# Instalamos los navegadores necesarios para Playwright
RUN playwright install chromium
RUN playwright install-deps chromium

# Comando para ejecutar el robot (asumiendo que crearás un script de inicio llamado main.py)
CMD ["python", "main_github.py"]
