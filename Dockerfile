# Sử dụng Python 3.11 slim image
FROM python:3.11-slim

# Thiết lập thư mục làm việc
WORKDIR /app

# Cài đặt các dependencies hệ thống cần thiết
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy file requirements trước để tận dụng Docker cache
COPY requirements.txt .

# Cài đặt Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ source code vào container
COPY . .

# Tạo thư mục uploads nếu chưa có
RUN mkdir -p static/uploads

# Tạo thư mục instance cho SQLite database
RUN mkdir -p instance

# Expose port 5000
EXPOSE 5000

# Biến môi trường
ENV FLASK_APP=app.py
ENV PYTHONUNBUFFERED=1

# Chạy ứng dụng
CMD ["python", "app.py"]
