# Usa uma versão leve do Python
FROM python:3.10-slim

# Define a pasta de trabalho lá no servidor do CPD
WORKDIR /app

# Copia os requisitos e instala
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia todo o seu código para dentro da caixa
COPY . .

# Comando para rodar o motor Raichu
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]