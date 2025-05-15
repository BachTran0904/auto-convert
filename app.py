from fastapi import FastAPI, UploadFile, File, HTTPException
import logging
import os
import tempfile
import subprocess
from fastapi.responses import JSONResponse, FileResponse
from pathlib import Path
from typing import List
import uvicorn

app = FastAPI()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


@app.post("/upload")
async def upload_and_process(files: List[UploadFile] = File(...), form: UploadFile = File(...)):
    temp_file_paths = []    
    temp_file_form_path = None
    #attribute_json = os.path.abspath('attribute.json')  # Use absolute path

    try:
        # Process all uploaded files
        for file in files:
            suffix = os.path.splitext(file.filename)[1] or '.bin'
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
                content = await file.read()
                temp_file.write(content)
                temp_file_paths.append(temp_file.name)
        
        # Process the form file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file_form:
            content = await form.read()
            temp_file_form.write(content)
            temp_file_form_path = temp_file_form.name

        if len(temp_file_paths) < 1:
            raise HTTPException(status_code=400, detail="At least one data file and one form file are required")
                
        # Build the command - note the order of arguments must match what mapping.py expects
        command = ["python", "mapping.py"] + temp_file_paths + [temp_file_form_path]
        logger.info(f"Running command: {' '.join(command)}")
        
        result = subprocess.run(
            command,
            capture_output=True,
            text=True
        )
        
        # Clean up the temporary input files
        for path in temp_file_paths:
            if os.path.exists(path):
                os.unlink(path)        

        if result.returncode != 0:
            logger.error(f"Script failed with error: {result.stderr}")
            raise HTTPException(status_code=400, detail=result.stderr)
            
        # Return the processed file for download
        return FileResponse(
            temp_file_form_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="processed_file.xlsx"
        )
        
    except Exception as e:
        # Clean up if something went wrong
        for path in temp_file_paths:
            if os.path.exists(path):
                os.unlink(path)
        if temp_file_form_path and os.path.exists(temp_file_form_path):
            os.unlink(temp_file_form_path)
        logger.error(f"Error in processing: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/get")
async def getOutput():
    output = "Form.xlsx"
    return output


# Add this block to run the application
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run(app, host="0.0.0.0", port=port)
