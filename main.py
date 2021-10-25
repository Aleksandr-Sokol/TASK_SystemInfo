import wmi
import socket
import winapps

file_name = 'test_name.txt'
all_informations = []
ip = input('Input ip:')
if ip:
    username = input('Input username:')
    password = input('Input password:')
    computer = wmi.WMI(ip, user=username, password=password)
else:
    computer = wmi.WMI()
computer_info = computer.Win32_ComputerSystem()[0]
os_info = computer.Win32_OperatingSystem()[0]
proc_info = computer.Win32_Processor()[0]
gpu_info = computer.Win32_VideoController()[0]
gpu_ram = int(gpu_info.AdapterRAM) / 1073741824  # B to GB
gpu_date = ''
try:
    if gpu_info.DriverDate:
        gpu_date = f'{gpu_info.DriverDate[:4]}-{gpu_info.DriverDate[4:6]}-{gpu_info.DriverDate[6:8]}'
except:
    all_informations.append(f'Error in gpu_info DriverDate')

os_name = os_info.Name.encode(encoding='ascii', errors='ignore').decode().split('|')[0].strip()
os_version = ' '.join([os_info.Version, os_info.BuildNumber])
system_ram = float(os_info.TotalVisibleMemorySize) / 1048576  # KB to GB
discs = computer.Win32_LogicalDisk()
discs_info = []
try:
    for disc in discs:
        size = int(disc.Size) / 1073741824  # B to GB
        free_size = int(disc.FreeSpace) / 1073741824  # B to GB
        discs_info.append(f"{disc.Caption[:-1]}({disc.Name[:-1]}, {disc.FileSystem}): {disc.Description}, "
                          f"size/ free: {size:.1f} GB/ {free_size:.1f} GB\n")
except:
    all_informations.append(f'Error in discs info')

try:
    hostname = socket.gethostname()
    IP = socket.gethostbyname(hostname)
    all_informations.append(f'Name PC: {computer_info.Name}, Hostname: {hostname}  IP: {IP}\n')
except:
    all_informations.append(f'Error hostname or IP info')

all_informations.append(f'Manufacturer: {computer_info.Manufacturer}, model: {computer_info.Model}\n')
all_informations.append(f'Domain: {computer_info.Domain}, user: {computer_info.UserName}\n')
all_informations.append(f'OS: {os_name}, v: {os_version}, {computer_info.SystemType} ({os_info.OSArchitecture}), '
      f'serial number: {os_info.SerialNumber}\n')
all_informations.append(f'CPU: {proc_info.Name} , {proc_info.NumberOfCores} cores\n')
all_informations.append(f'RAM: {system_ram:.2f} GB\n')
all_informations.append(f'Graphics Card: {gpu_info.Name}, RAM {gpu_ram:.1f} GB, driver {gpu_info.DriverVersion}'
                        f'/ {gpu_date}\n')
all_informations.append('Discs:\n')
for disc in discs_info:
    all_informations.append(disc)
all_informations.append('Applications:\n')
for i, a in enumerate(winapps.list_installed()):
    try:
        version = f' v.{a.version}' if a.version is not None else ''
        install_location = f'{a.install_location}' if a.install_location is not None else ''
        install_date = f'({a.install_date})' if a.install_date is not None else ''
        uninstall = f'uninstall: {a.uninstall_string}' if a.uninstall_string is not None else ''
        source = f'source: {a.install_source}' if a.install_source is not None else ''
        all_informations.append(f'{i}. {a.name} {version} {install_date} {install_location}| {source}| {uninstall}, '
                                f'{a.publisher}\n')
    except:
        all_informations.append(f'Error in application {i}')

try:
    file_name = f'{computer_info.Name}_IP_{IP}.txt'
    with open(file_name, 'w') as f:
        f.writelines(all_informations)
except:
    print(f'Error in application write to file')
