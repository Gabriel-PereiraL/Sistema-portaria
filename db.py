"""Database connection configuration loaded from the environment."""
import os
import mysql.connector

DB_CONFIG = {
    "host": os.environ.get("PORTARIA_DB_HOST", "localhost"),
    "user": os.environ["PORTARIA_DB_USER"],
    "password": os.environ["PORTARIA_DB_PASSWORD"],
    "database": os.environ.get("PORTARIA_DB_NAME", "portaria"),
    "port": int(os.environ.get("PORTARIA_DB_PORT", "3306")),
}

def get_connection():
    """Create a MySQL connection from environment-backed configuration."""
    return mysql.connector.connect(**DB_CONFIG)
