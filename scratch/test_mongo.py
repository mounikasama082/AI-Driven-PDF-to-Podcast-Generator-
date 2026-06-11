import os
from pymongo import MongoClient
from pymongo.server_api import ServerApi
from dotenv import load_dotenv

load_dotenv()
MONGO_URI = os.environ.get("MONGO_URI", "")

if not MONGO_URI:
    print("MONGO_URI not found in .env")
else:
    try:
        client = MongoClient(MONGO_URI, server_api=ServerApi("1"), serverSelectionTimeoutMS=5000)
        client.admin.command("ping")
        print("Successfully connected to MongoDB")
    except Exception as e:
        print(f"Failed to connect to MongoDB: {e}")
