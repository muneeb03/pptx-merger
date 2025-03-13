import os
from comtypes import client
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tqdm import tqdm

def convert_ppt_to_pdf(ppt_path):
    try:
        powerpoint = client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = True
        
        deck = powerpoint.Presentations.Open(str(ppt_path))
        pdf_path = str(ppt_path.with_suffix('.pdf'))
        
        # PDF format (17) is a constant in PowerPoint
        deck.SaveAs(pdf_path, 17)
        deck.Close()
        powerpoint.Quit()
        
        return True, ppt_path
    except Exception as e:
        return False, f"Error converting {ppt_path}: {str(e)}"

def bulk_convert(input_dir):
    input_path = Path(input_dir)
    ppt_files = list(input_path.glob("*.ppt*"))
    
    if not ppt_files:
        print("No PowerPoint files found in the specified directory.")
        return
    
    print(f"Found {len(ppt_files)} PowerPoint files.")
    
    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(tqdm(
            executor.map(convert_ppt_to_pdf, ppt_files),
            total=len(ppt_files),
            desc="Converting files"
        ))
    
    # Process results
    successful = [r[1] for r in results if r[0]]
    failed = [r[1] for r in results if not r[0]]
    
    print(f"\nConversion complete:")
    print(f"Successfully converted: {len(successful)} files")
    print(f"Failed conversions: {len(failed)} files")
    
    if failed:
        print("\nFailed conversions:")
        for error in failed:
            print(error)

if __name__ == "__main__":
    input_directory = input("Enter the directory path containing PPT files: ")
    bulk_convert(input_directory)