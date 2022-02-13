import xlrd
import sys
import struct

FORAMT = "{0},{1},{2},{}"


def read_xls(path):

    xf = xlrd.open_workbook(path)

    if xf is None:
        return

    sheet1 = xf.sheet_by_index(0)

    rows = sheet1.nrows
    cols = sheet1.ncols

    print(rows)
    print(cols)

    folders = []

    groupfile = open("gtmp","wb")

    folderfile = open("ftmp","wb")

    titlefile  = open("ttmp","w")

    outfile = open("out.txt","w")

    foldercount = 0
    groupcount  = 0

    for i in range(rows - 1):

        Type  = sheet1.cell(i + 1,0).value
        ID    = sheet1.cell(i + 1,1).value
        Alias = sheet1.cell(i + 1,2).value




        MCC   = sheet1.cell(i + 1,3).value
        MNC   = sheet1.cell(i + 1,4).value
        Addr  = sheet1.cell(i + 1,5).value
        Cou   = sheet1.cell(i + 1,6).value
        DmoFreq = sheet1.cell(i + 1,7).value

        gtype = 0

        print("%s,%s,%s,%s,%s,%s,%s,%s"%(Type,ID,Alias,MCC,MNC,Addr,Cou,DmoFreq))
        outfile.write("%s,%s,%s,%s,%s,%s,%s,%s\r\n"%(Type,ID,Alias,MCC,MNC,Addr,Cou,DmoFreq))

        if len(Type) <= 0:
            break





        if isinstance(ID,float):
            ID = int(ID)
        else:
            ID = 0

        if isinstance(MCC,float):
            MCC = int(MCC)
        else:
            MCC = 0

        if isinstance(MNC,float):
            MNC = int(MNC)
        else:
            MNC = 0

        if isinstance(Addr,float):
            Addr = int(Addr)
        else:
            Addr = 0

        if isinstance(Cou,float):
            Cou = int(Cou)
        else:
            Cou = 0

        if isinstance(DmoFreq,float):
            DmoFreq = int(DmoFreq * 1000000)
        else:
            DmoFreq = 0

        if isinstance(Type,str):

            if "Folders" in Type:

                folder = {"id":ID,"name":Alias}

                if folder not in folders:
                    folders.append(folder)
                    foldercount += 1
                    fold = struct.pack("h16s",ID,Alias.encode("utf-8"))
                    folderfile.write(fold)
                    continue


            if "TMO" in Type:

                gtype = 0

            elif "DMO" in Type:

                gtype = 1
            else:
                continue

            groupcount += 1

            grp = struct.pack("hh16shhIhI",gtype,ID,Alias.encode("utf-8"),MCC,MNC,Addr,Cou,DmoFreq)

            groupfile.write(grp)



    groupfile.close()
    folderfile.close()

    titlefile.write(str(foldercount))
    titlefile.write("\n")
    titlefile.write(str(groupcount))
    titlefile.close()

    return
if __name__ == "__main__":

    #p1 = sys.argv[1]


    #if(len(p1) > 0):

     #   read_xls(p1)

    read_xls("in.xlsx")


