import logging
import os

# * Check if logs file exists
logPath = os.path.join(os.getcwd(), "logs")
if not os.path.exists(logPath):
    os.mkdir(logPath)
# * Logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter("%(asctime)s:%(levelname)s:%(message)s")
file_handler = logging.FileHandler(os.path.join(logPath, "app.log"))
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
