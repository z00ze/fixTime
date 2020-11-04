# pip install openpyxl
import openpyxl as xl
import os, time

input_folder = os.path.join("input/")
output_folder = os.path.join("output/")



def fixTime(input_folder, file, col, row, debug):
    """ Reads excel file and fixes the time in a column
    """
    wb = xl.load_workbook(filename = input_folder + file)
    if debug: print("Fixing : " + input_folder + file)
    for sheet in wb:
        while sheet[col+str(row)].value != None:
            datas = str(sheet[col+str(row)].value).split('.')
            if len(datas) == 2:
                sheet[col+str(row)] = "00:%s:%s,000" % (datas[0].zfill(2), datas[1]) 
            row += 1
            
    """ Saving the modified xlsx to output folder
    """
    if debug: print("Saving : " + output_folder+file)
    wb.save(output_folder+file)

if __name__ == "__main__":
    """ Check if file exists in input folder and does not exist in output folder,
        then converts it to with fixTime function and save it to output folder.
    """
    # Start column
    col = 'A'
    # Start row
    row = 1
    while (True):
    
        if(len(os.listdir(input_folder)) > 0):
            
            inputs = os.listdir(input_folder)
            outputs = os.listdir(output_folder)
            
            for i in inputs:
                if not i in outputs:
                    fixTime(input_folder, i, col, row, True)

        time.sleep(1)
