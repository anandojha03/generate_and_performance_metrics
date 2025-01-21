import os
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime

def test_document_comparison():
    # Configuration
    BASE_URL = "http://127.0.0.1:8000"  # Adjust as needed
    ENDPOINT = "/compare-documents/"
    TOTAL_TESTS = 5
    
    # Setup directories
    form1_dir = Path("test_pages_jpg/form1")
    form2_dir = Path("test_pages_jpg/form2")
    
    # Initialize results storage
    successful_matches = []
    partial_matches = []
    no_matches = []
    complete_total_passed = 0
    partial_total_passed = 0
    no_total_passed = 0
    
    #Complete Match
    # Process each pair of files
    for i in range(1, TOTAL_TESTS + 1):
        filename = f"complete_match_{i}.jpg"
        file1_path = form1_dir / filename
        file2_path = form2_dir / filename
        
        # Skip if either file doesn't exist
        if not file1_path.exists() or not file2_path.exists():
            print(f"Skipping test {i}: One or both files missing")
            continue
        
        # Prepare files for upload
        files = {
            'file1': ('file1.jpg', open(file1_path, 'rb'), 'image/jpeg'),
            'file2': ('file2.jpg', open(file2_path, 'rb'), 'image/jpeg')
        }
        
        try:
            # Make API request
            print(f"Trying to hit fastAPI with {file1_path} and {file2_path}")
            response = requests.post(f"{BASE_URL}{ENDPOINT}", files=files)
            response.raise_for_status()
            print(f"Got some response")
            # Process response
            result = response.json()
            
            if result['comparison_result']['Status'] == "Complete Match":
                successful_matches.append({
                    'Test Number': i,
                    'File1': str(file1_path),
                    'File2': str(file2_path),
                    'Full Response': str(result)
                })
                complete_total_passed += 1
                
            print(f"Processed test {i}: {'Success' if result['comparison_result']['Status'] == 'Complete Match' else 'Failed'}")
            
        except Exception as e:
            print(f"Error processing test {i}: {str(e)}")
            
        finally:
            # Close file handles
            for f in files.values():
                f[1].close()

    #Partial Match
    # Process each pair of files
    for i in range(1, TOTAL_TESTS + 1):
        filename = f"partial_match_{i}.jpg"
        file1_path = form1_dir / filename
        file2_path = form2_dir / filename
        
        # Skip if either file doesn't exist
        if not file1_path.exists() or not file2_path.exists():
            print(f"Skipping test {i}: One or both files missing")
            continue
        
        # Prepare files for upload
        files = {
            'file1': ('file1.jpg', open(file1_path, 'rb'), 'image/jpeg'),
            'file2': ('file2.jpg', open(file2_path, 'rb'), 'image/jpeg')
        }
        
        try:
            # Make API request
            print(f"Trying to hit fastAPI with {file1_path} and {file2_path}")
            response = requests.post(f"{BASE_URL}{ENDPOINT}", files=files)
            response.raise_for_status()
            print(f"Got some response")
            # Process response
            result = response.json()
            
            if result['comparison_result']['Status'] == "Partial Match":
                partial_matches.append({
                    'Test Number': i,
                    'File1': str(file1_path),
                    'File2': str(file2_path),
                    'Full Response': str(result)
                })
                partial_total_passed += 1
                
            print(f"Processed test {i}: {'Success' if result['comparison_result']['Status'] == 'Partial Match' else 'Failed'}")
            
        except Exception as e:
            print(f"Error processing test {i}: {str(e)}")
            
        finally:
            # Close file handles
            for f in files.values():
                f[1].close()

    # Nothing Match
    # Process each pair of files
    for i in range(1, TOTAL_TESTS + 1):
        filename = f"no_match_{i}.jpg"
        file1_path = form1_dir / filename
        file2_path = form2_dir / filename
        
        # Skip if either file doesn't exist
        if not file1_path.exists() or not file2_path.exists():
            print(f"Skipping test {i}: One or both files missing")
            continue
        
        # Prepare files for upload
        files = {
            'file1': ('file1.jpg', open(file1_path, 'rb'), 'image/jpeg'),
            'file2': ('file2.jpg', open(file2_path, 'rb'), 'image/jpeg')
        }
        
        try:
            # Make API request
            print(f"Trying to hit fastAPI with {file1_path} and {file2_path}")
            response = requests.post(f"{BASE_URL}{ENDPOINT}", files=files)
            response.raise_for_status()
            print(f"Got some response")
            # Process response
            result = response.json()
            
            if result['comparison_result']['Status'] == "Nothing Match":
                no_matches.append({
                    'Test Number': i,
                    'File1': str(file1_path),
                    'File2': str(file2_path),
                    'Full Response': str(result)
                })
                no_total_passed += 1
                
            print(f"Processed test {i}: {'Success' if result['comparison_result']['Status'] == 'Nothing Match' else 'Failed'}")
            
        except Exception as e:
            print(f"Error processing test {i}: {str(e)}")
            
        finally:
            # Close file handles
            for f in files.values():
                f[1].close()
    
    # Create Excel report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"comparison_results_{timestamp}.xlsx"
    
    # Create DataFrame for successful matches
    df_complete_matches = pd.DataFrame(successful_matches)
    df_partial_matches = pd.DataFrame(partial_matches)
    df_no_matches = pd.DataFrame(no_matches)
    
    # Create Excel writer object
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write successful matches
        if successful_matches:
            df_complete_matches.to_excel(writer, sheet_name='Successful Matches', index=False)
        
        if partial_matches:
            df_partial_matches.to_excel(writer, sheet_name='Partial Matches', index=False)
        
        if no_matches:
            df_no_matches.to_excel(writer, sheet_name='Nothing Matches', index=False)
        
        # Write summary
        complete_match_summary_data = {
            'Metric': ['Total Tests', 'Successful Matches', 'Success Rate'],
            'Value': [
                TOTAL_TESTS,
                complete_total_passed,
                f"{(complete_total_passed/TOTAL_TESTS)*100:.2f}%"
            ]
        }
        pd.DataFrame(complete_match_summary_data).to_excel(writer, sheet_name='Complete Match Summary', index=False)

        partial_match_summary_data = {
            'Metric': ['Total Tests', 'Successful Matches', 'Success Rate'],
            'Value': [
                TOTAL_TESTS,
                partial_total_passed,
                f"{(partial_total_passed/TOTAL_TESTS)*100:.2f}%"
            ]
        }
        pd.DataFrame(complete_match_summary_data).to_excel(writer, sheet_name='Partial Match Summary', index=False)

        complete_match_summary_data = {
            'Metric': ['Total Tests', 'Successful Matches', 'Success Rate'],
            'Value': [
                TOTAL_TESTS,
                no_total_passed,
                f"{(no_total_passed/TOTAL_TESTS)*100:.2f}%"
            ]
        }
        pd.DataFrame(complete_match_summary_data).to_excel(writer, sheet_name='Nothing Match Summary', index=False)
    
    print(f"\nTest Summary:")
    print(f"Total tests: {TOTAL_TESTS}")
    print(f"Successful matches: {complete_total_passed}")
    print(f"Success rate: {(complete_total_passed/TOTAL_TESTS)*100:.2f}%")
    print(f"\nResults saved to: {output_file}")

if __name__ == "__main__":
    test_document_comparison()