FROM python:3.11-slim
 
# Install system dependencies

RUN apt-get update && apt-get install -y \

    graphviz \
&& rm -rf /var/lib/apt/lists/*
 
WORKDIR /app
 
COPY . /app
 
RUN pip install --upgrade pip

RUN pip install -r requirements.txt
 
EXPOSE 8501
 
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
 
