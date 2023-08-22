###################################################################################################################################
###         GENERATOR BARKODU                                                                                                   ###
###         v. 0.0.1                                                                                                            ###
###         17/8/2023                                                                                                           ###
###                                                                                                                             ###
###         Popis:                                                                                                              ###
###                 1, Načte vstup.xlsx Excel soubor s kodama CODE128 pod sebou ve sloupci "A"                                  ###
###                 2, Z načtených hodnot vygeneruje CODE128 PNG obrázky a uloží je do adresáře ./obrazky_code_128              ###
###                 3, Obrázky postupně v kládá do Excel souboru vystup.xlsx                                                    ###
###################################################################################################################################

# MODULY
import openpyxl 
import os
import barcode
from barcode.writer import ImageWriter
import PIL
from PIL import Image
# KONEC - MODULY


### POCATECNI PROMENE ###
# cesta_desktop - kam se maji ulozit vygenerovane obrazky barkodů 
cesta_desktop_pro_puvodni_velikost = os.environ['USERPROFILE'] + "\\Desktop\\BARCODE-GENERATOR\\generator-barkody-puvodni-velikost\\"
cesta_desktop_pro_zmenenou_velikost = os.environ['USERPROFILE'] + "\\Desktop\\BARCODE-GENERATOR\\generator-barkody-zmenena-velikost\\"
# Jestli se má generovat text pod obrázkem barkodu
generovat_text_pod_barkodem = True
### KONEC - POCATECNI PROMENE ###

### GLOBALNI PROMENE ###

# Funkce pro načtení všech zaznamu v souboru 
def loadAllEnteriesFromExcell(cislo_sloupce):
    #Nacteni excelovskeho worbooku ze souboru 
    book = openpyxl.load_workbook('Otis CPN.xlsx')
    #Vybrani aktivniho prvniho listu z listu z workbooku  
    sheet = book.active

    #Vytvoreni listu pro ukladani hodnot
    zaznamy = []

    # Iterovani vsech radek v Excellovskem listu 
    
    pocet_zaznamu = 1
   
    for row in sheet:
        # Dostani prvni hodnoty ze sloupce (cislo_sloupce), pozor cislovani je od nuly
        zaznam = row[cislo_sloupce].value
        #Podminka pokud zde nic neni, nevypisuj to
        print(pocet_zaznamu)
        if(zaznam == None):
            continue    

        writeBarcodesIntoExcell(generateBarcodeWithNewDimensions(generateBarcodeFromString(zaznam)))

def generateBarcodeFromString(alfanum_code):
    
    sample_barcode = barcode.get('code128', alfanum_code, writer=ImageWriter())
    ### NASTAVENI-POPISKU
    barcode.base.Barcode.default_writer_options['write_text'] = generovat_text_pod_barkodem
    ### NASTAVENI-POPISKU - KONEC
    generated_path_plus_filename = sample_barcode.save('' + cesta_desktop_pro_puvodni_velikost + alfanum_code)
    generated_filename = (alfanum_code+".png")
    print('Vytvoren barcode názvem a obsahem souboru: ' + generated_path_plus_filename)
    return generated_filename
  


def generateBarcodeWithNewDimensions(original_image):
    ### NASTAVENI-VELIKOSTI-OBRAZKU
    to_be_resized = Image.open('' + cesta_desktop_pro_puvodni_velikost + original_image) # otevřít v PIL Image Objekt
    newSize = (600, 100) # Nastaveni novych rozmeru
    resized = to_be_resized.resize(newSize, resample=PIL.Image.NEAREST) #je možné vybrat více druhů resamplů, NEAREST vypadá nejlíp
    generated_path_plus_filename_resized = resized.save(cesta_desktop_pro_zmenenou_velikost + 'resized_'+ original_image) # ulozeni obrazku s novými rozmery
    print('Soubor se zmenenou velikosti: ' + cesta_desktop_pro_zmenenou_velikost + 'resized_' + original_image)
    returning_for_writeBarcodesIntoExcell = original_image
    return returning_for_writeBarcodesIntoExcell
    ### KONEC - NASTAVENI-VELIKOSTI-OBRAZKU

def writeBarcodesIntoExcell(resized_image_for_excell):
    cell_number = 1
    cell = 'A' + str(cell_number)

    workbook = openpyxl.Workbook()
    worksheet = workbook.worksheets[0]
    
    print("Zapis souboru do Excelu: " + 'resized_' + resized_image_for_excell)
    print()

    img = openpyxl.drawing.image.Image(cesta_desktop_pro_zmenenou_velikost + 'resized_'+ resized_image_for_excell)
    img.anchor = cell
    worksheet.add_image(img)


    workbook.save('out.xlsx')

    cell_number = cell_number + 5
    print(cell)


# Definovani main funkce
def main():

    loadAllEnteriesFromExcell(cislo_sloupce=8)
  

  
if __name__=="__main__":
    main()
  


 
