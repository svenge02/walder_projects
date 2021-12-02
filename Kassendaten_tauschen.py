# Dieses Programm tauscht einzelne Daten aus dem Kassenprogramm aus.
# Die alten werden gesichert und ih ein Backup-Verzeichnis geschrieben.
# Die auszutauschenden Dateien werden in das auf dem Laufwerk G:\Controlling\Python\Kassendaten kopiert
# Wichitg ist, dass auf dem PC wo diese Python-Scripts ausgeführt werden SMB1 aktiviwert ist, sonst
# funktioniert das nicht.
# sg/23.07.2021


import os
import shutil as sh
import time


# filialen = {'101', '102', '103', '105', '107', '1142', '117', '1172', '120', '1202', '126', '127', '210',
#            '215', '216', '221', '224', '228', '233', '234', '240', '244', '406', '408', '411', '782', '793',
#            '790', '7902', '992', '993', '994'}

filialen = ['782']

src_files = r'G:\Controlling\Python\Kassendaten'  # Hier liegen die benötigen Files
back_verz = r'L:\BACKUP_2108'  # In dieses Verzeichnis werden die auszutauschenden Dateien gesichert
base_verz = os.path.join("L:", os.sep)

notw_files = os.listdir(src_files)  # Erstellt eine Liste der auszutauschenden Files
print(notw_files)

if os.path.exists(base_verz):
    os.system(r'net use L: /d /n')

for filiale in filialen:
    os.system(r'net use L: \\winpos' + filiale + '\c$\dtspos\DAT\FORMULAR etos /user:kasse')
    if not os.path.exists(back_verz):
        os.makedirs(back_verz)
    for file in notw_files:
        backup_file = os.path.join('L:' + os.sep + file)
        neues_file = os.path.join(src_files + os.sep + file)
        sh.copy(backup_file, back_verz)
        os.remove(backup_file)
        sh.copy(neues_file, base_verz)

    print(f'Fililae {filiale} ist erledigt')
    os.system(r'net use L: /d /n')


