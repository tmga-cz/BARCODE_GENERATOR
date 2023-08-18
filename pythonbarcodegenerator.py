#########################################################################################################
###         GENERATOR BARKODU
###         v. 0.0.1
###         17/8/2023
###
###         Popis: 
###                 1, Načte vstup.xlsx Excel soubor s kodama CODE128 pod sebou ve sloupci "A"
###                 2, Z načtených hodnot vygeneruje CODE128 PNG obrázky a uloží je do adresáře ./obrazky_code_128
###                 3, Obrázky postupně v kládá do Excel souboru vystup.xlsx
#########################################################################################################
import openpyxl 
from openpyxl import load_workbook

import barcode
from barcode.writer import ImageWriter


import PIL
from PIL import Image


# Funkce pro načtení všech zaznamu v souboru 
def load_all_enteries_in_excell():
    #Nacteni excelovskeho worbooku ze souboru 
    book = load_workbook('Otis CPN.xlsx')
    #Vybrani aktivniho prvniho listu z listu z workbooku  
    sheet = book.active

    #Vytvoreni listu pro ukladani hodnot
    zaznamy = []

    # Iterovani vsech radek v Excellovskem listu 
    for row in sheet:
        # Get the value of the first cell
        # in the row (the "Name" cell)
        zaznam = row[9].value
        # Add the value to the list
       
        if(zaznam == None):
            continue
            
        zaznamy.append(zaznam)

    #Vypsani vsech zaznamu v listu
    print(zaznamy)



load_all_enteries_in_excell()



code = 'Tomas'
sample_barcode = barcode.get('code128', code, writer=ImageWriter())

### NASTAVENI-POPISKU
barcode.base.Barcode.default_writer_options['write_text'] = True
### NASTAVENI-POPISKU - KONEC

generated_filename = sample_barcode.save('barcode2')
print('Generated Code 128 barcode image file name: ' + generated_filename)

### NASTAVENI-VELIKOSTI-OBRAZKU
to_be_resized = Image.open('barcode2.png') # open in a PIL Image object
newSize = (400, 100) # new size will be 500 by 300 pixels, for example
resized = to_be_resized.resize(newSize, resample=PIL.Image.NEAREST) # you can choose other :resample: values to get different quality/speed results
resized.save('filename_resized.png') # save the resized image
### KONEC - NASTAVENI-VELIKOSTI-OBRAZKU
  
wb = openpyxl.Workbook()
ws = wb.worksheets[0]

# ws.append([10, 2010, "Geeks", 4, "life"])
img = openpyxl.drawing.image.Image('barcode2.png')
  
img.anchor = 'A8'

ws.add_image(img)
wb.save('out.xlsx')



