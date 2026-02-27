import winreg
import binascii

def get_monitor_names_from_registry():
    monitors = []
    try:
        # Enum DISPLAY
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Enum\DISPLAY")
        for i in range(winreg.QueryInfoKey(key)[0]):
            dev_id = winreg.EnumKey(key, i)
            dev_key = winreg.OpenKey(key, dev_id)
            for j in range(winreg.QueryInfoKey(dev_key)[0]):
                inst_id = winreg.EnumKey(dev_key, j)
                inst_key = winreg.OpenKey(dev_key, inst_id)
                try:
                    params_key = winreg.OpenKey(inst_key, "Device Parameters")
                    edid, _ = winreg.QueryValueEx(params_key, "EDID")
                    
                    # Parse EDID for 0xFC (Monitor Name)
                    name = None
                    for k in range(4):
                        offset = 54 + k * 18
                        if offset + 18 <= len(edid):
                            block = edid[offset:offset+18]
                            if block[0:3] == b'\x00\x00\x00' and block[3] == 0xFC:
                                name_bytes = block[5:]
                                # Find newline 0x0A
                                idx = name_bytes.find(b'\x0a')
                                if idx != -1:
                                    name_bytes = name_bytes[:idx]
                                name = name_bytes.decode('ascii', errors='ignore').strip()
                                break
                    if name:
                        monitors.append((dev_id, name))
                except OSError:
                    pass
    except Exception as e:
        print("Registry Error:", e)
    return monitors

print("Monitors from Registry EDID:", get_monitor_names_from_registry())
