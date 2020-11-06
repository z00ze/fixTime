# pip install openpyxl
import openpyxl as xl
import os, time

input_folder = os.path.join("input/")
output_folder = os.path.join("output/")



def fixTime(input_folder, output_folder, name, time_col, header_col, end_row, debug):
    """ Reads excel file and fixes the time in a column
    """
    # Start row
    row = 5
    wb = xl.load_workbook(filename = input_folder + name, data_only=True)
    if debug: print("Fixing : " + input_folder + name)
    for sheet in wb:
        while row < end_row:

            if len(str(sheet[header_col+str(row)].value)) == 0 or str(sheet[header_col+str(row)].value) == "None":
                row += 1
                continue
            else:
                if len(str(sheet[time_col+str(row)].value)) == 0 or str(sheet[time_col+str(row)].value) == "None":
                    header = str(sheet[header_col+str(row)].value).replace(" ", "_").replace(".", "_").replace("?", "_").replace(":", "_")
                    print(header)
                    file = open(output_folder+header+".srt","w")
                    index = 1
                    row += 2
                    
                    while sheet[time_col+str(row)].value != None:
                        if len(str(sheet[time_col+str(row)].value).split(" ")) > 1:
                            sheet[time_col+str(row)].value = str(sheet[time_col+str(row)].value).split(" ")[1]
                        datas = str(sheet[time_col+str(row)].value).split(':')
                                 
                        if len(datas) == 3:
                            start_time = "00:%s:%s,000" % (datas[0].zfill(2), datas[1].zfill(2))
                            if int(datas[1]) + 3 > 59:
                                datas[0] = str(int(datas[0]) + 1)
                                datas[1] = str((int(datas[1]) + 3) % 60)
                            else:
                                datas[1] = str(int(datas[1])+3)

                            end_time = "00:%s:%s,000" % (datas[0].zfill(2), datas[1].zfill(2))
                            file.write(str(index)+"\n")
                            file.write(start_time + " --> " + end_time +"\n")
                            file.write(str(sheet[header_col+str(row)].value)+"\n\n")
                            index += 1
                        row += 2
                    file.close()

        

if __name__ == "__main__":
    """ Check if file exists in input folder and does not exist in output folder,
        then converts it to with fixTime function and save it to output folder.
    """
    # Start column
    time_col = 'B'
    # Header column
    header_col = 'C'
    
    # End row
    end_row = 1000

    
    if(len(os.listdir(input_folder)) > 0):

        inputs = os.listdir(input_folder)

        for i in inputs:

            if i.split(".")[-1:][0] == "xlsx":
                
                fixTime(input_folder, output_folder, i, time_col, header_col, end_row, True)


