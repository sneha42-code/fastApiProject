import shutil

def save_upload_file(upload_file, destination: str):
    """
    Save an uploaded file to the specified destination
    """
    with open(destination, "wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)
