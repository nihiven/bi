## informatics magic box

## required libraries
## watchdog: https://github.com/gorakhargosh/watchdog
## pytest: https://github.com/pytest-dev/pytest
## XLRD: https://github.com/python-excel/xlrd
## Pillow: https://github.com/python-pillow/Pillow

## settings.lua
settings = {}
settings['path'] = '.\\testDir'

# filetype handler classes
fileTypeHandlers = {}


## watchdog 
import sys
import time
import logging
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# excel
import xlrd

#pillow 
from PIL import Image


####
## filetype handler classes
####
class clsExcelFileTypeHandler():
    name = "excel"
    description =  "Microsoft Excel Spreadsheet"
    fileTypes = { "xls", "xlsx" }

    def test(self, fileName):
        try:
            book = xlrd.open_workbook(fileName)
            return { "success": True , "name": clsExcelFileTypeHandler.name, "format": book.biff_version }
        except xlrd.biffh.XLRDError as e:
            return { "success": False }

# end clsExcelFileTypeHandler


class clsImageFileTypeHandler():
    name = "image"
    description = "Various Image Formats"
    fileTypes = { "jpg", "png", "gif" } # TODO: not sure about this

    def test(self, fileName):
        try:
            im = Image.open(fileName)
            return { "success": True, "name": clsImageFileTypeHandler.name, "format": im.format }
        except IOError as e:
            return { "success": False }

# end clsImageFileTypeHandler


####################

####
## magic box functions
####
def registerFileTypeHandler(fileTypeHandler):
    fileTypeHandlers[fileTypeHandler.name] = fileTypeHandler()

def processFile(fileName):
    fileType = getFileType(fileName)

    if (fileType == False):
        print "Unknown Filetype"
    else:
        print "", "File:", fileName
        print "", "File Type:", fileType["name"], fileType["format"]

    print ""
# end processFile()


def getFileType(fileName):
    for handler in fileTypeHandlers:
        result = fileTypeHandlers[handler].test(fileName)
        if (result['success']):
            return result

    return False
# end getFileType()


####
## watchdog event handlers
####
class clsMyEventHandler(FileSystemEventHandler):
    #event properties: is_directory event_type src_path

   # def on_any_event(self, event):
   #     if (event.is_directory == False):
   #         print "", event

    def on_modified(self, event):
        if (event.is_directory == False):
            processFile(event.src_path)
    # end on_modified()

    def on_moved(self, event):
        if (event.is_directory == False):
            processFile(event.src_path)
    # end on_moved()

    def on_created(self, event):
        if (event.is_directory == False):
            processFile(event.src_path)
    # end on_created()

## clsMyEventHandler


####
## main
####

# register our filetype handlers
registerFileTypeHandler(clsExcelFileTypeHandler)
registerFileTypeHandler(clsImageFileTypeHandler)

print "Available Filetype Handlers"
for handler in fileTypeHandlers:
    print "", fileTypeHandlers[handler].name, ":", fileTypeHandlers[handler].description

print ""


# run it
path = settings['path']
event_handler = clsMyEventHandler()
observer = Observer()
observer.schedule(event_handler, path, recursive=True)
observer.start()

print "Monitoring", path

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer.join()
