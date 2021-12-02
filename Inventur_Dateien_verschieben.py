"""
Dieses Programm verschiebt die Inventur-Datein vom FilialPC in ein Archiv auf dem FilialPC und auf diesen PC

SG/11.06.2021
"""
 # Notwendige Module
import os
import shutil
import re
import ctypes
import sys

# Runs the program as Administrator
# ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)


def walk_error(error):
    print(error.filename)

FIL = '21'   # nur 2stellig, wird in IP-Adressen eingefügt
TYP = '100'
INV_JAHR = '2021'


# In diesem Verzeichnis müssen sich die Kontoauszüge befinden
source = r'\\192.168.'+FIL+'.100\mde$'
dest_arch_fil = os.path.join(r'\\192.168.' + FIL + '.' + TYP + '\mde$'  + os.sep, 'Archiv' + os.sep, INV_JAHR)
dest_zentrale = os.path.join(r'C:\etos\mde\check')

# if not os.path.exists(dest_arch_fil):
#    os.makedirs(dest_arch_fil)

# Die Datein in diesem Verzeichnis werden gelesen ...
for folderName, subfolers, filenames in os.walk(source, onerror=walk_error):

    for filename in filenames:

        filename_von = os.path.join(source + os.sep, filename)
        filename_nach_archiv = os.path.join(dest_arch_fil + os.sep, filename)
        filename_nach_zentrale = os.path.join(dest_zentrale + os.sep, filename )
        if '.nor' in filename:
            print(filename)
            shutil.copy(filename_von, filename_nach_archiv)
            shutil.move(filename_von, filename_nach_zentrale)

    break # Dieser Break verhindert, dass auch die Subfolder gelesen werden.

print('Dateien erfolgreich verschoben')