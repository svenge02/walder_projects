'''
Da im Verzeichnis DTSPOS\DAT Dateien überschrieben werden, soll das ganze Verzeichnis
vorgängig kopiert werden. Dabei muss auch sichergestellt werden, dass dieses Backup auf
allen Kassen durchgeführt wird
'''

import os
import shutil as sh

# Die letzten 2 Nummern der Filiale, da diese für die IP-Adress benötigt werden (ausser 101)
filialen = ['101', '2', '3', '105']

# os.system(r'net use L: \\192.168.2.101\c$\dtspos etos /user:kasse')

src = r'\\192.168.10.101\c$:\dtspos\DAT'
dest = r'\\192.168.10.101\c$:\dtspos\DAT_2107'

# if not os.path.exists(dest):
#   os.makedirs(dest)

if not os.path.exists(dest):
    sh.copytree(src, dest)

# os.system(r'net use L: /d /y')


