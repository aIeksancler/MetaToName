from pathlib import Path
import glob, os, re

from win32com.propsys import propsys, pscon
from datetime import datetime


video_suffixes = ['.mp4', '.avi', '.mov', '.m4v', '.mpg']
photo_suffixes = ['.jpg', '.jpeg', '.raw', '.png', '.arw']


dir_path = r'.\**\*'

file_list = glob.glob(dir_path, recursive=True)
for file in file_list:

    date = None
    year = None
    make = None
    model = None
    suffixes = None
    file_path = None
    ## get all suffixes
    suffixes = Path(file).suffixes

    ## get file directory
    file_path = Path(os.path.join(file)).parent

    try:
        print(f'Working with: {file}')
        if str(suffixes[-1].lower()) in video_suffixes or str(suffixes[-1].lower()) in photo_suffixes:
            #print(f'{suffixes[-1]} is in {video_suffixes} or {photo_suffixes}')
            properties = propsys.SHGetPropertyStoreFromParsingName(os.path.abspath(file))
            date = properties.GetValue(pscon.PKEY_Photo_DateTaken).GetValue()
            if date != None:
                make = properties.GetValue(pscon.PKEY_Photo_CameraManufacturer).GetValue()
                model = properties.GetValue(pscon.PKEY_Photo_CameraModel).GetValue()
                #print(f'Photo date: {date} {make} {model}')
            else:
                date = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()
                if date != None:
                    pass

            ## construct file name and path
            properties = None
            new_name = ''
            if date != None:
                year = date.strftime('%Y')

                ## folder sorting
                if not os.path.exists(year):
                    os.mkdir(year)
                
                date = date.strftime('%Y%m%d_%H%M%S%z')
                new_name = date

                if make != None:
                    new_name += '_' + re.sub(r'[^_\-A-Za-z0-9]+', '', make)
                    
                if model != None:
                    new_name += '_' + re.sub(r'[^_\-A-Za-z0-9]+', '', model)



                ## check if file exists and change file name if neccessary
                copy_number = 0
                while copy_number < 20:

                    full_path = os.path.join(file_path)

                    if not Path(os.path.join(file)).parent == Path(year):
                        full_path = os.path.join(full_path, year)
                        print(f'Not in a folder: {full_path}')
                    
                    if copy_number > 0:
                        full_path = os.path.join(full_path, (new_name.upper() + '_(' + str(copy_number) + ')'))
                    else:
                        full_path = os.path.join(full_path, new_name.upper())

                    for suffix in suffixes:
                        if str(suffix) != '.s':
                            full_path += str(suffix).lower()

                    #if Path(file).stem == Path(full_path).stem:
                    if os.path.abspath(file) == os.path.abspath(full_path):
                        print(f'{file} matches new formatting. Skipping.')
                        break
                    
                    elif not os.path.exists(full_path):
                        ## rename file
                        print('OLD: ' + file)
                        print('NEW: ' + full_path)
                        os.rename(file, full_path)
                        
                        if os.path.exists(file + '.pp3'):
                            print(f'Renaming .pp3 files')
                            print(f'OLD: {file}.pp3')
                            print(f'NEW: {full_path}.pp3')
                            os.rename(file + '.pp3', full_path + '.pp3')
                            
                        #print(f'New name: {new_name}')
                        break

                    print(f'Exisiting file: {full_path}')
                    copy_number += 1
            else:
                print(f'No metadata in file: {file}')
        else:
            print(f'Unsupported file: {suffixes[-1]}')

            
    except Exception as e:
        print(f'Error on file: {file}')
        print(e)
        
    properties = None

input("Press Enter to exit...")
