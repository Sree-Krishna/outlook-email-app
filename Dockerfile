# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /app

# Copy the current directory contents into the container
ADD . /app

# Install dependencies
RUN pip install -r requirements.txt

# Expose the Flask port
EXPOSE 8000

# Run the Flask app
CMD ["python", "app.py"]
