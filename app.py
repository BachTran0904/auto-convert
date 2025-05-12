from fastapi import FastAPI, UploadFile, File, HTTPException
import logging
import os
import tempfile
import subprocess
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse

app = FastAPI()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@app.post("/upload")
async def mapping(file: UploadFile = File(...)):
    # Create a temporary file to save the uploaded content
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        # Save the uploaded file to the temporary location
        content = await file.read()
        temp_file.write(content)
        temp_file_path = temp_file.name
    
    # Define other file paths (consider using pathlib for cross-platform compatibility)
    formatted = 'C:/Github/auto-convert/Form.xlsx'  # Using forward slashes works on Windows too
    attribute_json = 'C:/Github/auto-convert/atribute.json'
    
    try:
        # Run the mapping.py script and capture output
        result = subprocess.run(
            ["python", "mapping.py", temp_file_path, formatted, attribute_json],
            capture_output=True,
            text=True
        )
        
        # Clean up the temporary file
        os.unlink(temp_file_path)
        
        # Return the stdout if successful, or stderr if failed
        if result.returncode == 0:
            return {"status": "success", "output": result.stdout}
        else:
            return {"status": "error", "error": result.stderr}
            
    except Exception as e:
        # Clean up the temporary file if something went wrong
        if os.path.exists(temp_file_path):
            os.unlink(temp_file_path)
        return {"status": "error", "error": str(e)}