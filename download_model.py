import os
import urllib.request
import zipfile
import shutil

MODEL_URL = "https://alphacephei.com/vosk/models/vosk-model-small-en-us-0.15.zip"
ZIP_PATH = "vosk_model.zip"
TARGET_DIR = "model"

def download_and_extract():
    if os.path.exists(TARGET_DIR) and os.listdir(TARGET_DIR):
        print(f"'{TARGET_DIR}' folder already exists and is not empty. Skipping download.")
        return

    print(f"Downloading lightweight VOSK model from {MODEL_URL} ...")
    urllib.request.urlretrieve(MODEL_URL, ZIP_PATH)
    print("Download completed. Extracting...")

    with zipfile.ZipFile(ZIP_PATH, 'r') as zip_ref:
        zip_ref.extractall("temp_model")

    # Find the extracted folder inside temp_model
    extracted_items = os.listdir("temp_model")
    extracted_folder = os.path.join("temp_model", extracted_items[0])

    if os.path.exists(TARGET_DIR):
        shutil.rmtree(TARGET_DIR)
    shutil.move(extracted_folder, TARGET_DIR)

    # Cleanup
    if os.path.exists("temp_model"):
        shutil.rmtree("temp_model")
    if os.path.exists(ZIP_PATH):
        os.remove(ZIP_PATH)

    print(f"VOSK model successfully extracted to '{TARGET_DIR}'.")

if __name__ == "__main__":
    download_and_extract()
