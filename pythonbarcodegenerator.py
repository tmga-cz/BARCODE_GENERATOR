###################################################################################################################################
###         GENERATOR BARKODU                                                                                                   ###
###         v. 0.0.2                                                                                                            ###
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
cesta_desktop = os.environ['USERPROFILE'] + "\\Desktop\\generator-barkody\\"
# Jestli se má generovat text pod obrázkem barkodu
generovat_text_pod_barkodem = False
### KONEC - POCATECNI PROMENE ###


# Funkce pro načtení všech zaznamu v souboru 
def loadAllEnteriesInExcell(cislo_sloupce):
    #Nacteni excelovskeho worbooku ze souboru 
    book = openpyxl.load_workbook('Otis CPN.xlsx')
    #Vybrani aktivniho prvniho listu z listu z workbooku  
    sheet = book.active

    #Vytvoreni listu pro ukladani hodnot
    zaznamy = []

    # Iterovani vsech radek v Excellovskem listu 
    for row in sheet:
        # Dostani prvni hodnoty ze sloupce (cislo_sloupce), pozor cislovani je od nuly
        zaznam = row[cislo_sloupce].value
        #Podminka pokud zde nic neni, nevypisuj to
        if(zaznam == None):
            continue    

        generateBarcodeFromString(zaznam)
        # Pridani hodnoty do listu     
        zaznamy.append(zaznam)

    #Vypsani vsech zaznamu v listu
    print(zaznamy)

def generateBarcodeFromString(alfanum_code):
    sample_barcode = barcode.get('code128', alfanum_code, writer=ImageWriter())
    ### NASTAVENI-POPISKU
    barcode.base.Barcode.default_writer_options['write_text'] = generovat_text_pod_barkodem
    ### NASTAVENI-POPISKU - KONEC
    generated_filename = sample_barcode.save('' + cesta_desktop + alfanum_code)
    print('Byl vygenerován soubor s Code128 a názvem a obsahem souboru: ' + generated_filename)


def generateBarcodeWithNewDimensions():
    ### NASTAVENI-VELIKOSTI-OBRAZKU
    to_be_resized = Image.open('barcode2.png') # otevřít v PIL Image Objekt
    newSize = (400, 100) # Nastaveni novych rozmeru
    resized = to_be_resized.resize(newSize, resample=PIL.Image.NEAREST) #je možné vybrat více druhů resamplů, NEAREST vypadá nejlíp
    resized.save('filename_resized.png') # ulozeni obrazku s novými rozmery
    ### KONEC - NASTAVENI-VELIKOSTI-OBRAZKU



# Definovani main funkce
def main():
    loadAllEnteriesInExcell(cislo_sloupce=8)

if __name__=="__main__":
    main()
  


 
