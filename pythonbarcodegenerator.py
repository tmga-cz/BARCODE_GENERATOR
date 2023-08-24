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
import os
# KONEC - MODULY


### POCATECNI PROMENE ###
# cesta_desktop - kam se maji ulozit vygenerovane obrazky barkodů 
cesta_desktop_pro_puvodni_velikost = os.environ['USERPROFILE'] + "\\Desktop\\BARCODE-GENERATOR\\generator-barkody-puvodni-velikost\\"
cesta_desktop_pro_zmenenou_velikost = os.environ['USERPROFILE'] + "\\Desktop\\BARCODE-GENERATOR\\generator-barkody-zmenena-velikost\\"
# Jestli se má generovat text pod obrázkem barkodu
generovat_text_pod_barkodem = True

### GLOBALNI PROMENE ###
pocet_zaznamu = 1
excel_souradnice_y = 1
###
workbook = openpyxl.Workbook()
worksheet = workbook.worksheets[0]

# Funkce pro mazani starych zaznamu
def delete_files_in_directory_and_subdirectories(directory_path):
    try:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                os.remove(file_path)
            print("Všechny stare podadresare a soubory v " + " - " + directory_path + " byly smazány")
    except OSError:
        print("Při mazání souborů  a podadresářů se vyskytly chyby")


# Usage
directory_path = '/path/to/directory'
delete_files_in_directory_and_subdirectories(directory_path)

# Funkce pro načtení všech zaznamu v souboru 
def loadAllEnteriesFromExcel(cislo_sloupce):
    global pocet_zaznamu
    global excel_souradnice_y

    #Nacteni excelovskeho worbooku ze souboru 
    book = openpyxl.load_workbook('Otis CPN.xlsx')
    #Vybrani aktivniho prvniho listu z listu z workbooku  
    sheet = book.active

    #Vytvoreni listu pro ukladani hodnot
    zaznamy = []

    # Iterovani vsech radek v Excelovskem listu
   
    for row in sheet:
        # Dostani prvni hodnoty ze sloupce (cislo_sloupce), pozor cislovani je od nuly
        zaznam = row[cislo_sloupce].value
        #Podminka pokud zde nic neni, nevypisuj to
        #print(pocet_zaznamu)
        if(zaznam == None):
            continue    
        writeBarcodesIntoExcel(generateBarcodeWithNewDimensions(generateBarcodeFromString(zaznam)))
      
        pocet_zaznamu += 1
        excel_souradnice_y += 5

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
    returning_for_writeBarcodesIntoExcel = original_image
    return returning_for_writeBarcodesIntoExcel
    ### KONEC - NASTAVENI-VELIKOSTI-OBRAZKU

def writeBarcodesIntoExcel(resized_image_for_excel):
    global excel_souradnice
    excel_souradnice_y
    excel_souradnice = 'A' + str(excel_souradnice_y)
    print("Zapis na souradnice ---> " + excel_souradnice + " do Excelu: " + 'resized_' + resized_image_for_excel+"\n")
    img = openpyxl.drawing.image.Image(cesta_desktop_pro_zmenenou_velikost + 'resized_'+ resized_image_for_excel)
    img.anchor = excel_souradnice
    worksheet.add_image(img)
 


# Definovani main funkce
def main():
    delete_files_in_directory_and_subdirectories(cesta_desktop_pro_puvodni_velikost)
    delete_files_in_directory_and_subdirectories(cesta_desktop_pro_zmenenou_velikost)
    print()
    loadAllEnteriesFromExcel(cislo_sloupce=8)
    workbook.save('out.xlsx')
    print("Konec")

    


if __name__=="__main__":
    main()
    
  


 
