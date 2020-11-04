# pip install openpyxl
import openpyxl as xl
import os, time

input_folder = os.path.join("input/")
output_folder = os.path.join("output/")



def fixTime(input_folder, file, col, row, end_row, debug):
    """ Reads excel file and fixes the time in a column
    """
    wb = xl.load_workbook(filename = input_folder + file, data_only=True)
    if debug: print("Fixing : " + input_folder + file)
    for sheet in wb:
        while row < end_row:
            if len(str(sheet[col+str(row)].value)) == 0:
                row += 1
                continue
            if len(str(sheet[col+str(row)].value).split(" ")) > 1:
                sheet[col+str(row)].value = str(sheet[col+str(row)].value).split(" ")[1]
            datas = str(sheet[col+str(row)].value).split(':')
            
            if len(datas) == 3:
                
                sheet[col+str(row)] = "00:%s:%s,000" % (datas[0].zfill(2), datas[1].zfill(2))
                
                if int(datas[1]) + 3 > 59:
                    datas[0] = str(int(datas[1]) + 1)
                    datas[1] = str((int(datas[1]) + 3) % 60)
                else:
                    datas[1] = str(int(datas[1])+3)
                    
                sheet[col+str(row+1)] = "00:%s:%s,000" % (datas[0].zfill(2), datas[1].zfill(2))
            row += 2
            
    """ Saving the modified xlsx to output folder
    """
    if debug: print("Saving : " + output_folder+file)
    wb.save(output_folder+file)

if __name__ == "__main__":
    """ Check if file exists in input folder and does not exist in output folder,
        then converts it to with fixTime function and save it to output folder.
    """
    # Start column
    col = 'B'
    # Start row
    row = 7
    # End row
    end_row = 1000
    while (True):
    
        if(len(os.listdir(input_folder)) > 0):
            
            inputs = os.listdir(input_folder)
            outputs = os.listdir(output_folder)
            
            for i in inputs:
                if not i in outputs:
                    fixTime(input_folder, i, col, row, end_row, True)

        time.sleep(1)
