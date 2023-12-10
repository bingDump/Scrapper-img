import requests # request img from web
import shutil # save img locally
import xlwings as xw
import timeit


start = timeit.default_timer() # mencatat waktu mulai

    # make sure the excell are in the same folder 
    # due to the img download will be save right here
    # where path that you open folder in Text Editor (VsCode)
destination_folder = "Dum Folder/SQUARE/"

workbook = xw.Book("Dum Folder/DOWNLOAD IMAGE AOS-28127.xlsx")
Wsheet = workbook.sheets("SQUARE")
sku_list = Wsheet.range("A2:A22").value
Link1_list = Wsheet.range("B2:B22").value
Link2_list = Wsheet.range("C2:C22").value
Link3_list = Wsheet.range("D2:D22").value
Link4_list = Wsheet.range("E2:E22").value
Link5_list = Wsheet.range("F2:F22").value

    # dictionary with list comprehension of lists above. if there is more or less than 4 links, you must change the variable above and object of dictionary below! 

data_sku_link = {
    sku_list[i]: [Link1_list[i], Link2_list[i], Link3_list[i], Link4_list[i], Link5_list[i]] for i in range(len(sku_list))} # type: ignore



# sku_dt = data_sku_link.keys()

    # list for file that success or failed to download
fail_dw = []
succeed_dw = []

for sku in data_sku_link:
    for link in data_sku_link[sku]:
        # implement the get() method from the requests module to retrieve the image. The method will take in two parameters, the url variable you created earlier, and stream: True by adding this second argument in guarantees no interruptions will occur when the method is running.
        res = requests.get(link, stream = True) # type: ignore
        file_name = f"{sku}-{data_sku_link[sku].index(link)+1}.jpg"
        if res.status_code == 200:
            # "wb" means open/create file for writing and open in binary mode
            with open(f"{destination_folder}{file_name}",'wb') as f:
                #The copyfileobj() method to write your image as the file name, builds the file locally in binary-write mode and saves it locally with shutil
                shutil.copyfileobj(res.raw, f)
            print(f'{len(succeed_dw)+1}. Img Downloaded: ',file_name)
            succeed_dw.append(file_name)
        else:
            print('Image Couldn\'t be retrieved: ',file_name)
            fail_dw.append(file_name)
            continue

    #count SKU that downloaded, "dummyset" is only dummy value for set declaration
sku_succeed_dw = {"dummyset"}
for sku_dw in succeed_dw:
    new_sku = sku_dw[:-6] #Slicing string SKU "-1.jpg"
    sku_succeed_dw.add(new_sku)

print(f" {len(succeed_dw)}\tfiles Sucessfully Downloaded")
print(f" {len(sku_succeed_dw)-1}\tSKU")
print(f" {len(fail_dw)}\tfiles Download Failed")

    # output list img that can't be downloaded
if len(fail_dw) > 0 :
    for failed in fail_dw:
        print(failed)

stop = timeit.default_timer() # catat waktu selesai
runtime = stop - start # lama eksekusi dalam satuan detik
print("Runtime: ",runtime,"sec")