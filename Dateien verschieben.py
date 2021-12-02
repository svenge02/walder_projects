# Dieses Probramm verschiebt die heruntergeladenen UBS-PDF Dokumente in
# die dafür vorgesehenen Ordner
# SG/20.04.2021

 # Notwendige Module
import os, shutil, re


def walk_error(error):
    print(error.filename)


# In diesem Verzeichnis müssen sich die Kontoauszüge befinden
verz = r'G:\Buchhaltung Lehrlinge\Kontoauszüge\UBS'

# Die Datein in diesem Verzeichnis werden gelesen ...
for folderName, subfolers, filenames in os.walk(verz, onerror=walk_error):

    for filename in filenames:
        print(filename)

        name_as_list = filename.split('_')

        print(name_as_list)

        if len(name_as_list) == 3:
            konto = name_as_list[0]
            bez = name_as_list[1]
            jahr = name_as_list[2][0:4]
            monat = name_as_list[2][4:6]
        elif len(name_as_list) == 4:
            konto = name_as_list[0]
            bez = name_as_list[2]
            jahr = name_as_list[3][0:4]
            monat = name_as_list[3][4:6]
        else:
            break

        newpath = os.path.join('G:' + os.sep, 'Buchhaltung Lehrlinge' + os.sep, 'Kontoauszüge' \
                               + os.sep, 'UBS' + os.sep, konto + os.sep, jahr + os.sep, monat \
                               + os.sep, bez)


        if not os.path.exists(newpath):
            os.makedirs(newpath)

        filename_von = os.path.join(verz + os.sep, filename)
        filename_nach = os.path.join(newpath + os.sep, filename)

        shutil.move(filename_von, filename_nach)

    break # Dieser Break verhindert, dass auch die Subfolder gelesen werden.

print('Dateien erfolgreich verschoben')