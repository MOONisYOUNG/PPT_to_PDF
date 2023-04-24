import sys
import os
import comtypes.client
from tqdm.notebook import tqdm

def ppt_to_pdf(input_path, output_path):
    
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    
    input_file_paths = os.listdir(input_path)
    
    for input_file_name in tqdm(input_file_paths):

        if not input_file_name.lower().endswith((".ppt", ".pptx")):
            continue

        input_file_path = os.path.join(input_path, input_file_name)
        print(input_file_name)

        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = True
        slides = powerpoint.Presentations.Open(input_file_path)

        file_name = os.path.splitext(input_file_name)[0]
        output_file_path = os.path.join(output_path, file_name + ".pdf")

        slides.SaveAs(output_file_path, FileFormat=32)
        slides.Close()
        
    print('\n--> Finish Work!!')