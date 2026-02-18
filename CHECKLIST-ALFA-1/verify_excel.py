import os
import sys
import pandas as pd
import datetime

# Add parent directory to path to import checklist_recondicionado
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from checklist_recondicionado import guardar_em_excel, formatar_excel, EXCEL_FILE

def test_excel_export():
    print("Testing Excel Export...")
    
    # Mock data
    usuario = "Tester"
    compra_num = "TEST-123"
    sys_info = {
        'modelo': 'Test PC',
        'serial': 'SN123456',
        'cpu': 'Intel Test',
        'ram': '16 GB',
        'disk': '512 GB SSD',
        'gpu': 'NVIDIA Test'
    }
    
    # New structure of tests
    testes = {
        "Teclado": True,
        "Ecrã": True,
        "Touch Screen": False,
        "Wifi": True,
        "LAN": False,
        "Webcam": True,
        "Microfone": True,
        "USB": True,
        "Portas de Vídeo": True
    }
    
    danos = "Teste de verificação automatica."
    
    # Call the function
    # Note: guardar_em_excel internally uses EXCEL_FILE which is defined in checklist_recondicionado.
    # We rely on it to use the path defined there or we might need to patch it if we want a temporary file.
    # For now, let's just run it and see if it adds a row to the actual file (or creates it).
    # Ideally we should use a temp file, but modifying the global var in imported module is tricky without refactor.
    # Let's try to mock the global EXCEL_FILE if possible or just use a test file name if I can change it.
    
    # Actually, I can just change variable in the module
    import checklist_recondicionado
    checklist_recondicionado.EXCEL_FILE = "test_registos.xlsx"
    
    if os.path.exists("test_registos.xlsx"):
        os.remove("test_registos.xlsx")
        
    success = guardar_em_excel(usuario, compra_num, sys_info, testes, danos)
    
    if success:
        print("Export function returned True.")
        
        # Verify columns
        df = pd.read_excel("test_registos.xlsx")
        expected_columns = [
            "Data", "Técnico", "Nº Compra", "Modelo", "Serial", "CPU", "RAM", "Disco", "GPU",
            "Teclado", "Ecrã", "Wifi", "LAN", "Touch Screen", "Webcam", "Microfone", "USB", 
            "Portas de Vídeo", "Notas"
        ]
        
        # Check if all expected columns are present
        missing = [col for col in expected_columns if col not in df.columns]
        
        if not missing:
            print("All expected columns are present.")
            print(f"Columns found: {list(df.columns)}")
            
            # Additional check: "Touch Screen" column exists
            if "Touch Screen" in df.columns:
                 print("SUCCESS: Touch Screen column found.")
            else:
                 print("FAILURE: Touch Screen column MISSING.")
                 
        else:
            print(f"FAILURE: Missing columns: {missing}")
            print(f"Columns found: {list(df.columns)}")
            
        # Clean up
        # os.remove("test_registos.xlsx") 
    else:
        print("Export function returned False.")

if __name__ == "__main__":
    test_excel_export()
