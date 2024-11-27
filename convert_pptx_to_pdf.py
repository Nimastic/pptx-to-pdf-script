import os
import subprocess

def convert_pptx_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file in os.listdir(input_folder):
        if file.endswith(".pptx"):
            input_path = os.path.join(input_folder, file)
            output_path = os.path.join(output_folder, file.replace(".pptx", ".pdf"))
            subprocess.run([
                "/Applications/LibreOffice.app/Contents/MacOS/soffice", "--headless",
                "--convert-to", "pdf", "--outdir", output_folder, input_path
            ])
    print("Conversion complete!")


input_dir = "/Users/jerielchan/Downloads/CS2106/pptx"
output_dir = "/Users/jerielchan/Downloads/CS2106/pdf"
convert_pptx_to_pdf(input_dir, output_dir)
