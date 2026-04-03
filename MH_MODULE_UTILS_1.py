import subprocess
from deep_translator import GoogleTranslator
from collections import defaultdict, Counter
from typing import List, Dict
import re
import unicodedata
from unidecode import unidecode
from PIL import Image, ImageDraw, ImageFont
watermarkdir = "/Users/bernardconti/LOCAL_TEMP/watermark"
#============================================================================================================= 
#============================================================================================================= 
def group_similar_texts(
    texts: List[str],
) -> Dict[str, List[str]]:
    """
    Groups texts that become identical after normalization.
    Good for typos in casing, extra spaces, punctuation.
    """
    groups = defaultdict(list)
    clean_texts = []
    les_keys =[]
    
    for text in texts:
        key = text[0]
        key = key.lower()
        key.lstrip()
        key.rstrip()
        key = re.sub(r'\d+', '', key)       # Remove all digits from the string
        key = re.sub(r'[^\w\s]', '', key)   # remove punctuation 
        key = ' '.join(key.split())         # normalize spaces
        clean_texts.append([key,text[1],len(key)])
    
    #sorted_keys = sorted(keys, lambda) #, reverse=True
    clean_texts = sorted(clean_texts, key=lambda col: (col[0],col[2]))

    for item in clean_texts:
        print(item)
        key = item[0]
        if key:

            for idx,item2 in enumerate(clean_texts):
                key2 = item2[0]
                if key in key2 :clean_texts[idx][0] = key

    for item in clean_texts:
        groups[item[0]].append(item[1])
        les_keys.append(item[0])

    return dict(groups),les_keys,clean_texts
#=============================================================================================================
def strip_accents(text):

    try:
        text = unidecode(text, 'utf-8')
    except NameError: # unicode is a default on python 3 
        pass

    text = unicodedata.normalize('NFD', text)\
           .encode('ascii', 'ignore')\
           .decode("utf-8")

    return str(text)
#=============================================================================================================       
def traduction_date(texte,*args):
    ishort = None
    for valeur in args:
        if valeur:
            if isinstance(valeur, str): valeur =valeur.lower()
            if valeur == "short" : ishort = valeur  

    texte = texte.replace("jan","janvier")
    texte = texte.replace("feb","février")
    texte = texte.replace("mar","mars")
    texte = texte.replace("apr","avril")
    texte = texte.replace("may","mai")
    texte = texte.replace("jun","juin")
    texte = texte.replace("jul","juillet")
    texte = texte.replace("aug","août")
    texte = texte.replace("sep","septembre")
    texte = texte.replace("oct","octobre")
    texte = texte.replace("nov","novembre")
    texte = texte.replace("dec","décembre")

    #texte = texte.replace(" ","/")  
    texte = texte.replace("about","env") 
    texte = texte.replace("after","après") 
    texte = texte.replace("depuis abt ","") 
    return texte
#=============================================================================================================
def traduction_mois_numero(texte):
    texte = texte.lower()
    texte = texte.replace("janvier","01")
    texte = texte.replace("février","02")
    texte = texte.replace("mars","03")
    texte = texte.replace("avril","04")
    texte = texte.replace("mai","05")
    texte = texte.replace("juin","06")
    texte = texte.replace("juillet","07")
    texte = texte.replace("août","08")
    texte = texte.replace("septembre","09")
    texte = texte.replace("octobre","10")
    texte = texte.replace("novembre","11")
    texte = texte.replace("décembre","12")
    texte = texte.replace(" ","/")    
    return texte
#=============================================================================================================
def traduction(le_texte):
#-------------------------------------------------------------------------------------------------------------  
    if le_texte:
        translated = le_texte
        if len(le_texte) < 200:
            try: translated = GoogleTranslator(source='en', target='fr').translate(le_texte)
            except:next
#-------------------------------------------------------------------------------------------------------------  
    return translated
#=============================================================================================================
def list_unique_colomn(la_list,n_col):
    list_unique = []
    [list_unique.append(item[n_col]) for item in la_list if item[n_col] not in list_unique]
    return list_unique
#=============================================================================================================
def watermark(image_path,le_texte): 
# --------------------------------------------------------------   
#Import required Image library
    #Create an Image Object from an Image
    if le_texte : 
        im = Image.open(image_path)
        width, height = im.size

        final_im = Image.new('RGB', (width, int(height*1.11)))
        watermark_im = Image.new('RGB', (width, int(height*0.11)),color = (256, 256, 256))
        draw = ImageDraw.Draw(watermark_im)
        fsize = height * 0.1
        font = ImageFont.truetype("Monaco.ttf",fsize)

        # calculate the x,y coordinates of the text
        x = width/10
        y =0

        position = (x, y)
        left, top, right, bottom = draw.textbbox(position, le_texte, font=font)
        #draw.rectangle((left-10, top-10, right+10, bottom+10), fill="white")
        #draw.rectangle((left, top, right, bottom), fill="white")
        draw.text(position, le_texte, font=font, fill="black")
        #draw.text(position, le_texte, font=font, fill="white")

        #im.show()
        w_file_name = watermarkdir+"/w_"+image_path.split("/")[-1]

        #Save watermarked image
        final_im.paste(im, (0,0))
        final_im.paste(watermark_im, (0,height))

        final_im.save(w_file_name)
    else: w_file_name = image_path

    return w_file_name
#=============================================================================================================
def excute_cmd(command): 
# --------------------------------------------------------------   
    command_list = command.split(" ")
    #print(command_list)
    result = subprocess.run(command_list, capture_output=True, text=True)
    if result.stdout : print ("resultat = " + result.stdout)
    if result.stderr : print("erreur = " + result.stderr)
# --------------------------------------------------------------
    return(result.stdout,result.stderr)
#=============================================================================================================