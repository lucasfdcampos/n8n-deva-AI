FROM n8nio/n8n:latest

# Mudar para root para instalar pacotes
USER root

# Instalar Python e ferramentas necessárias (Alpine Linux)
RUN apk add --no-cache \
    python3 \
    py3-pip

# Instalar bibliotecas Python para processar XLSX
RUN pip3 install --break-system-packages --no-cache-dir \
    openpyxl \
    pandas \
    xlrd \
    numpy

# Voltar para o usuário node (segurança)
USER node
