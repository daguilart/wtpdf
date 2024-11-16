import os
import comtypes.client

def initialize_word():
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    return word

def convert_to_pdf(word, input_file, output_file):
    # Open the input file
    doc = word.Documents.Open(input_file)
    # Export it as a PDF file
    doc.SaveAs(output_file, FileFormat=17)  # 17 is the file format code for PDF
    doc.Close()

def bulk_convert_word_to_pdf(input_directory):
    output_directory = os.path.join(input_directory, 'PDF')
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    word = initialize_word()
    
    for root, _, files in os.walk(input_directory):
        for file in files:
            if file.endswith('.docx') or file.endswith('.doc'):
                input_file = os.path.join(root, file)
                relative_path = os.path.relpath(root, input_directory)
                output_folder = os.path.join(output_directory, relative_path)
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)
                
                output_file = os.path.join(output_folder, os.path.splitext(file)[0] + '.pdf')
                convert_to_pdf(word, input_file, output_file)
                print(f"Converted {input_file} to {output_file}")

    word.Quit()

if __name__ == "__main__":
    input_directory = input("Enter the directory to search for Word files: ")
    bulk_convert_word_to_pdf(input_directory)
