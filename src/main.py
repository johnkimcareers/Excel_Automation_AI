from dotenv import load_dotenv
load_dotenv()
import os
import glob
from assignment import Assignment

# Assuming 'Assignment' class is defined as provided earlier and has the necessary methods
def process_files(base_directory):
    # List all subdirectories in the base directory
    subdirectories = [f.path for f in os.scandir(base_directory) if f.is_dir()]
    example_assignment = Assignment(r'/Users/exampleUser/Desktop/ExampleFolder/example.xlsx')

    for directory in subdirectories:
        # Use glob to find all matching .docx and .xlsx files
        docx_files = glob.glob(os.path.join(directory, 'example_*.docx'))
        xlsx_files = glob.glob(os.path.join(directory, 'example_*rubric.xlsx'))

        # Sort the files to ensure matching files are aligned
        docx_files.sort()
        xlsx_files.sort()

        # Process each pair of files
        for docx_file, xlsx_file in zip(docx_files, xlsx_files):
            # Extract the base name without extension to ensure they match
            base_name_docx = os.path.splitext(os.path.basename(docx_file))[0]
            base_name_xlsx = os.path.splitext(os.path.basename(xlsx_file))[0]

            # Check if the base names match (without the 'rubric' part for the Excel files)
            if base_name_docx == base_name_xlsx.replace('rubric', ''):
                # Process the files
                example_assignment.parse_assignment(docx_file)
                example_assignment.parse_marked_rubric(xlsx_file)
            else:
                print(f"Warning: Mismatched files found: {docx_file} and {xlsx_file}")
    # write to file
    example_assignment.create_example_doc()

# Usage
base_directory = r'/Users/exampleUser/Desktop/ExampleFolder/'
process_files(base_directory)
