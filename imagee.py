import os
import cv2
import shutil
import zipfile
import tempfile
from docx import Document
import streamlit as st
from io import BytesIO

def extract_and_process_images(docx_file):
    temp_directory = tempfile.mkdtemp()
    extracted_images_folder = os.path.join(temp_directory, "extracted_images")
    processed_images_folder = os.path.join(temp_directory, "processed_images")
    os.makedirs(extracted_images_folder, exist_ok=True)
    os.makedirs(processed_images_folder, exist_ok=True)

    st.write("Temporary directories created.")

 
    try:
        word_document = Document(docx_file)
        for image_index, relationship in enumerate(word_document.part.rels.values()):
            if "image" in relationship.target_ref:
                image_data = relationship.target_part.blob
                extracted_image_path = os.path.join(extracted_images_folder, f"image_{image_index + 1}.png")
                with open(extracted_image_path, "wb") as image_file:
                    image_file.write(image_data)
                st.write(f"Extracted image saved: {extracted_image_path}")
    except Exception as e:
        st.error(f"Failed to extract images from the DOCX file: {e}")
        shutil.rmtree(temp_directory)
        return None

    for image_name in os.listdir(extracted_images_folder):
        original_image_path = os.path.join(extracted_images_folder, image_name)
        processed_image_path = os.path.join(processed_images_folder, image_name)

      
        image = cv2.imread(original_image_path, cv2.IMREAD_GRAYSCALE)
        if image is None:
            st.error(f"Could not read the image: {original_image_path}")
            continue

        
        inverted_image = 255 - image

        
        cv2.imwrite(processed_image_path, inverted_image)
        st.write(f"Processed image saved: {processed_image_path}")

   
    zip_file_path = os.path.join(temp_directory, "processed_images.zip")
    with zipfile.ZipFile(zip_file_path, "w") as zip_file:
        for root, _, files in os.walk(processed_images_folder):
            for file in files:
                zip_file.write(os.path.join(root, file), arcname=file)
    st.write(f"ZIP file created: {zip_file_path}")


    shutil.rmtree(extracted_images_folder)
    shutil.rmtree(processed_images_folder)

    return zip_file_path

st.title("DOCX Image Processor")
st.write("Convert Terminal programming or any pics from dark to white to avoid getting scolded by shopkeeper:))))")
st.write("PS: It happened to some of my friends")

uploaded_docx_file = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded_docx_file is not None:
    st.write("Valid DOCX file uploaded. Processing...")
    
  
    docx_file = BytesIO(uploaded_docx_file.read())
    
  
    zip_file_path = extract_and_process_images(docx_file)
    
    if zip_file_path:
        with open(zip_file_path, "rb") as zip_file:
            st.download_button(
                label="Download Processed Images ZIP",
                data=zip_file,
                file_name="processed_images.zip",
                mime="application/zip"
            )
        st.success("Processing complete!")
    else:
        st.error("An error occurred during processing.")
