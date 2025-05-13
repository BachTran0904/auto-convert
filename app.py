from fastapi import FastAPI, UploadFile, File, HTTPException
import logging
import os
import tempfile
import subprocess
from fastapi.responses import JSONResponse, FileResponse
from pathlib import Path
import uvicorn  # Add this import

app = FastAPI()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@app.post("/upload")
async def upload_and_process(file: UploadFile = File(...), form: UploadFile = File(...)):
    # Create a temporary file to save the uploaded content
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        content = await file.read()
        temp_file.write(content)
        temp_file_path = temp_file.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file_form:
        content = await form.read()
        temp_file_form.write(content)
        temp_file_form_path = temp_file_form.name
    
    # Define other file paths
    attribute_json = 'C:/Github/auto-convert/atribute.json'
    
    try:
        # Run the mapping.py script
        result = subprocess.run(
            ["python", "mapping.py", temp_file_path, temp_file_form_path, attribute_json],
            capture_output=True,
            text=True
        )
        
        # Clean up the temporary input file
        os.unlink(temp_file_path)
        
        if result.returncode != 0:
            raise HTTPException(status_code=400, detail=result.stderr)
            
        # Return the processed file for download
        return FileResponse(
            temp_file_form_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="processed_file.xlsx"
        )
        
    except Exception as e:
        # Clean up if something went wrong
        if os.path.exists(temp_file_path):
            os.unlink(temp_file_path)
        raise HTTPException(status_code=500, detail=str(e))
    

@app.get("/get")
async def getOutput():
    output = "Form.xlsx"
    return output

# Add this block to run the application
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run(app, host="0.0.0.0", port=port)