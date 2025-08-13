import os
import win32com.client # Ensure you have pywin32 installed

def ppt_to_pdf(input_path, output_path):
    """
    Converts a PowerPoint file to a PDF file using win32com.
    """
    if not os.path.exists(input_path):
        print(f"Error: Input file not found at {input_path}")
        return

    ppSaveAsPDF = 32
    
    try:
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        presentation = powerpoint.Presentations.Open(input_path, ReadOnly=True)
        presentation.SaveAs(output_path, ppSaveAsPDF)
        presentation.Close()
        powerpoint.Quit()
        
        print(f"Successfully converted '{input_path}' to '{output_path}'")
        
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    # Use the file names directly since they're in the same folder
    input_file_name = "AgileVsWaterfall.pptx"
    output_file_name = "AgileVsWaterfall.pdf"
    
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Construct the full paths
    input_path = os.path.join(script_dir, input_file_name)
    output_path = os.path.join(script_dir, output_file_name)
    
    # Call the conversion function
    ppt_to_pdf(input_path, output_path)