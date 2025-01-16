import os
import subprocess

def convert_pptx_to_pdf(input_folder, output_folder):
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Loop through all .pptx files in the input folder
    for file in os.listdir(input_folder):
        if file.endswith(".pptx"):
            input_path = os.path.join(input_folder, file)
            output_path = os.path.join(output_folder, file.replace(".pptx", ".pdf"))
            try:
                # Run the LibreOffice command for Windows
                subprocess.run([
                    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
                    "--headless",
                    "--convert-to", "pdf", 
                    "--outdir", output_folder, 
                    input_path
                ], check=True)
                print(f"Converted: {file}")
            except subprocess.CalledProcessError as e:
                print(f"Error converting {file}: {e}")
            except FileNotFoundError:
                print("LibreOffice not found. Please check the path to soffice.exe.")
                return

    print("Conversion complete!")

# Define input and output directories
input_dir = r"C:\Users\jerie\Downloads\LectureNotes_export"
output_dir = r"C:\Users\jerie\Downloads\CS4231_Lectures"

# Convert files
convert_pptx_to_pdf(input_dir, output_dir)
