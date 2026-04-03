import os
import sqlite3

# file management
import glob
import os.path
from html.parser import HTMLParser

list_events = ["Décoration","Distinction","Degree","Diplôme",
                "Military Service","Award","Honors","Title","Titre",
                "Military Award","Military Enlistment","Residence","Separation",
                "Anoblissement","Nomination,Immigration","Association",
                "Illness","Comment","Marriage","","Custom event","Nationalité","Celebrity"
                ]
#"Language spoken",

#===========================================================================================
# GEDCOM
def clean_CONC(idx,les_lignes_r,file_output):
#-------------------------------------------------------------------------------------------------------------  
    #print(f"ligne      {idx} = {les_lignes_r[idx]}")

    if les_lignes_r[idx] == "0 TRLR":
        file_output.write(les_lignes_r[idx])
    else:
        les_mots = les_lignes_r[idx].split(" ")
        if les_mots[0].isdigit():

            next_les_mots = les_lignes_r[idx+1].split(" ")
            if not next_les_mots[0].isdigit():

                replace_next_ligne = les_lignes_r[idx][:-1] + les_lignes_r[idx+1]
                les_lignes_r[idx+1] = replace_next_ligne
                idx = clean_CONC(idx + 1,les_lignes_r,file_output)

            else: file_output.write(les_lignes_r[idx])
        else:
            print('clean_CONC ... Pas normal !')
            print(f"{idx} ... {les_lignes_r[idx]}")
    #---------------------------------------------------------------------------------------  
    return idx
#===========================================================================================
def clean_place(la_place):
    clean_place = None
    if la_place: 
        clean_place = la_place.lower()
        clean_place = clean_place.replace('(',',')
        clean_place = clean_place.replace('[',',')
        clean_place = clean_place.replace(')',',')
        clean_place = clean_place.replace(']',', ')
        clean_place = clean_place.replace('-',' ')
        clean_place = clean_place.replace(', ',',')
        clean_place = clean_place.replace(' sur ','/')
        clean_place = clean_place.replace(' arrondissement ','')
        temp2 = []
        for t in clean_place.split(","):
            t = t.lstrip()
            temp2.append(t.capitalize())
        if temp2 : clean_place = ", ".join(temp2)

    return clean_place
#===========================================================================================
def clean_date(la_date):
    clean_date = None
    if la_date:
        clean_date = la_date.replace("AND","-")
        clean_date = clean_date.replace("FROM AFT","à partir de")
        clean_date = clean_date.replace("AFT","après")
        clean_date = clean_date.replace("BET","")
        clean_date = clean_date.replace("FROM","")
        clean_date = clean_date.replace("TO","-")

        clean_date = clean_date.replace("JAN","janvier")
        clean_date = clean_date.replace("FEB","février")
        clean_date = clean_date.replace("MAR","mars")
        clean_date = clean_date.replace("APR","avril")
        clean_date = clean_date.replace("MAY","mai") 
        clean_date = clean_date.replace("JUN","juin")
        clean_date = clean_date.replace("JUL","juillet")
        clean_date = clean_date.replace("AUG","août")
        clean_date = clean_date.replace("SEP","septembre")
        clean_date = clean_date.replace("OCT","octobre")
        clean_date = clean_date.replace("NOV","novembre")
        clean_date = clean_date.replace("DEC","décembre")
        clean_date = clean_date.lstrip()
        clean_date = clean_date.rstrip()
    return clean_date
#===========================================================================================
def clean_type(le_type):
    if le_type:
        if      le_type == "Degree" : le_type = "Diplôme"
        elif    le_type == "Military Service" : le_type = "Activité militaire"
        elif    le_type == "Military Enlistment" : le_type = "Activité militaire"
        elif    le_type == "Award"  : le_type = "Distinction"
        elif    le_type == "Honors" : le_type = "Distinction"
        elif    le_type == "Military Award" : le_type = "Distinction militaire"
        elif    le_type == "Title" : le_type = "Titre"
        elif    le_type == "Anoblissement" : le_type = "Titre"
        elif    le_type == "Illness" : le_type = "Maladie"
        elif    le_type == "Comment" : le_type = "Divers"
        elif    le_type == "Marriage" : le_type == "Mariage"
        elif    le_type == "Separation" : le_type == "Séparation"
        elif    le_type == "Custom event" : le_type = "Divers"
    return le_type
#===========================================================================================
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
#===========================================================================================
def convert_note(note_html):
    isVerbose = False
#------------------------------------------------------------------------------------------------------------- 
    class MyHTMLParser(HTMLParser):
        def handle_starttag(self, tag, attrs):
            la_liste = ["starttag"]
            la_liste.append(tag)
            for attr in attrs: la_liste.append(attr)
            les_data.append(la_liste)
        def handle_endtag(self, tag):
            les_data.append(["endtag",tag])
        def handle_data(self, data):
            les_data.append(["data",data])  
#------------------------------------------------------------------------------------------------------------- 
    les_textes = []
    les_href = []
    le_texte = ""
    parser = MyHTMLParser()
    #-----------------------------------------------------------------------------------------------------
    # HTPL parsing 
    les_data = []
    parser.feed(note_html)  
    # Initialisation des flags
    is_linkname = False
    is_linkurl = False
    prefix = ""

    isHref = False


    # traitement des données parsées
    #---------------------------------------------------------------------------
    for la_data in les_data: 

        if isVerbose :  print("la_data=",la_data[0],la_data[1])            
        # starttag 
        #-----------------------------------------------------------------------
        if la_data[0] == "starttag":

            # p,figcaption, h2
            if la_data[1][0] == "p" or la_data[1] == "figcaption" or la_data[1] == "h2":
                if le_texte: 
                    les_textes.append(f'{prefix}{le_texte}')
                    le_texte = ""

            if la_data[1] == "a":  isHref = True
            # li
            if la_data[1] == "li": 
                if le_texte: 
                    les_textes.append(f'{prefix}{le_texte}')
                    le_texte = ""
                prefix = "• "

            if la_data[1] == "strong": le_texte = le_texte + "<strong>"

            # linkurl
            if la_data[1] == "linkurl"  : is_linkurl = True
            if la_data[1] == "linkname" : is_linkname = True

        # data
        #-----------------------------------------------------------------------
        if la_data[0] == "data":
            if isVerbose : print (la_data[1])
            if "http" in la_data[1]: les_href.append(la_data[1])
            else:
                temp_text = la_data[1]
                temp_text = temp_text.replace("§","")
                ts = temp_text.replace(" ","")
                if temp_text != "Web content link:" and ts and not is_linkurl and not is_linkname:
                    le_texte = le_texte + temp_text
        # endtags
        #-----------------------------------------------------------------------
        if la_data[0] == "endtag":
            if isVerbose : print (la_data[1])
            if la_data[1] == "li" : 
                if le_texte: 
                    les_textes.append(f'{prefix}{le_texte}')
                    le_texte = ""
                prefix = ""
            if la_data[1] == "linkurl" : is_linkurl = False
            if la_data[1] == "linkname" : is_linkname = False
            if la_data[1] == "a": 
                isHref = False
            if la_data[1] == "strong": le_texte = le_texte + "</strong>"
#-------------------------------------------------------------------------------------------  
    if le_texte: les_textes.append(le_texte)

    return les_textes,les_href
#===========================================================================================
def sqlite_gedcom2sql(dir):
#===========================================================================================
# DEBUT
#===========================================================================================
    files_in_dir = glob.glob(dir + "clean_*.ged")
    for file in files_in_dir:
        try: os.remove(file)
        except FileNotFoundError: print(f"File {file} is not present in the system.")

    files_in_dir = glob.glob(dir + "*.ged")
    if files_in_dir : 
        latest_file_in_dir = max(files_in_dir, key=os.path.getctime)
        gedcom_fichier = os.path.basename(latest_file_in_dir)
        #(file_name_noextention,extension) = os.path.splitext(file_name)
        full_gedcom_fichier= dir + gedcom_fichier
    else:
        print("pas de fichier .ged dans",dir)
        exit()

    database_file = "/Users/bernardconti/Downloads/"+ gedcom_fichier.replace(".ged",".db")
    if not os.path.isfile(database_file):
    #if True:
        print(120 * "=")
        print ("Génération de la base de donnée SQL à partir de " + gedcom_fichier)
        print(120 * "=")
        print("")
        #===========================================================================================
        # Clean Gedcom File
        #if gedcom_fichier.split("_")[0] !="clean":
        isClean = True
        if isClean:
            clean_gedcom_fichier_1 = "/Users/bernardconti/Downloads/"+"clean_1_" + gedcom_fichier
            clean_gedcom_fichier_2 = "/Users/bernardconti/Downloads/"+"clean_2_" + gedcom_fichier
            
        # first pass : retire les SOURCE
            file_input = open(full_gedcom_fichier, 'r', encoding='utf8',errors='ignore')
            file_output = open(clean_gedcom_fichier_1, 'w')

            les_lignes = file_input.readlines()
            file_input.close()
            is_write = True
            for idx, la_ligne in enumerate(les_lignes):
                les_mots = la_ligne.split(" ")
                if idx > 15:
                    if les_mots[0] == "1" :
                            if (les_mots[1] == "SOUR" 
                                or les_mots[1] == "PUBL"
                                or les_mots[1] == "EMAIL"
                                ) : is_write = False
                            else: is_write = True
            
                if is_write : file_output.write(la_ligne)

            file_output.close()
            #print("Pass 1 : fait")

        # second pass : corrige les CONC
            file_input = open(clean_gedcom_fichier_1, 'r', encoding='utf8',errors='ignore')
            file_output = open(clean_gedcom_fichier_2, 'w')

            les_lignes_r = file_input.readlines()

            idx = -1
            for item in les_lignes_r:
                idx = idx + 1
                
                if idx > 15: 
                    if idx < len(les_lignes_r):
                        idx = clean_CONC(idx,les_lignes_r,file_output)    
                else:
                    file_output.write(les_lignes_r[idx])

            file_output.close()
            #print("Pass 2 : fait")
            full_gedcom_fichier = clean_gedcom_fichier_2

        #===========================================================================================
        # Connect to the SQLite database (or create it if it doesn't exist)
        #===========================================================================================
        database_file = "/Users/bernardconti/Downloads/"+ gedcom_fichier.replace(".ged",".db")
        if not os.path.isfile(database_file):
            # Connect to the SQLite database (or create it if it doesn't exist)

            connection_obj = sqlite3.connect(database_file)
            print(f"Opened SQLite database {database_file} with version {sqlite3.sqlite_version} successfully.")
            sql = connection_obj.cursor()

            # SQL query to create the table
            creation_query = """
                CREATE TABLE INDI (
                    indi_id INTEGER PRIMARY KEY,
                    nom TEXT,
                    prenom TEXT,
                    prenoms TEXT,
                    surnom TEXT,
                    sexe CHAR(1),
                    bdate TEXT,
                    bplace TEXT,
                    isdead CHAR(1),
                    ddate TEXT,
                    dplace TEXT,
                    cause TEXT
                );"""
            sql.execute("DROP TABLE IF EXISTS INDI")
            sql.execute(creation_query)
            #---------------------------------------------------
            creation_query = 'CREATE TABLE FAMC (indi_id INTEGER,fam_id INTEGER,isAdopted TEXT);'
            sql.execute("DROP TABLE IF EXISTS FAMC")
            sql.execute(creation_query)
            #---------------------------------------------------
            creation_query = 'CREATE TABLE FAMS (indi_id INTEGER, fam_id INTEGER);'
            sql.execute("DROP TABLE IF EXISTS FAMS")
            sql.execute(creation_query)
            #-------------------------------------------------
            creation_query = 'CREATE TABLE HUSB (fam_id INTEGER,indi_id INTEGER);'
            sql.execute("DROP TABLE IF EXISTS HUSB")
            sql.execute(creation_query)
            #-------------------------------------------------
            creation_query = 'CREATE TABLE WIFE (fam_id INTEGER,indi_id INTEGER);'
            sql.execute("DROP TABLE IF EXISTS WIFE")
            sql.execute(creation_query)
            #-------------------------------------------------
            creation_query = 'CREATE TABLE CHIL (fam_id INTEGER,indi_id INTEGER);'
            sql.execute("DROP TABLE IF EXISTS CHIL")
            sql.execute(creation_query)
            #-------------------------------------------------
            creation_query = """ CREATE TABLE OBJE (indi_id INTEGER,
                                    form TEXT,
                                    url TEXT,
                                    title TEXT,
                                    date TEXT,
                                    place TEXT,
                                    personal CHAR(1));"""
            sql.execute("DROP TABLE IF EXISTS OBJE")
            sql.execute(creation_query)
            #-------------------------------------------------
            creation_query = 'CREATE TABLE INFO (indi_id INTEGER, even TEXT, type TEXT, description TEXT, date TEXT, place TEXT, note TEXT);'
            sql.execute("DROP TABLE IF EXISTS INFO")
            sql.execute(creation_query)
            #-------------------------------------------------
            creation_query = 'CREATE TABLE BIOS (indi_id INTEGER, note TEXT);'
            sql.execute("DROP TABLE IF EXISTS BIOS")
            sql.execute(creation_query)
            #-------------------------------------------------
            creation_query = 'CREATE TABLE HREF (indi_id INTEGER, href TEXT);'
            sql.execute("DROP TABLE IF EXISTS HREF")
            sql.execute(creation_query)

            #connection_obj.commit()
            #-------------------------------------------------
            file_input = open(clean_gedcom_fichier_2, 'r', encoding='utf8',errors='ignore')
            #========================================================================================
            #ZOB
            #=================================================================================================
            # First pass = INDI
            #=================================================================================================
            idx = 0
            idf = 0
            fam_id = None
            MH_indi = [None] * 12
            isFirst_indi = True
            isFirst_obje_file = True
            isFirst_even =  True
            isFirst_bio = True
            for ligne in file_input: 
                ligne= ligne.replace('\n', '')
                MH_records = ligne.split(" ")
                #=================================================================================================
                if MH_records[0] == "0" :
                #=================================================================================================
                    if MH_records[-1] == "INDI":
                        if isFirst_indi : isFirst_indi = False
                        else:
                            INDI_row = '''
                            INSERT INTO INDI (indi_id,nom,prenom,prenoms,surnom,sexe,bdate,bplace,isdead,ddate,dplace,cause) 
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?);
                            '''
                            INDI_data = (MH_indi[0],MH_indi[1],MH_indi[2],MH_indi[3],MH_indi[4],MH_indi[5],
                                         MH_indi[6],MH_indi[7],MH_indi[8],MH_indi[9],MH_indi[10],MH_indi[11]) 
                            sql.execute(INDI_row, INDI_data)
                            
                        idx = idx + 1
                        MH_indi = [None] * 12
                        indi_id = int(MH_records[1][2:-1])
                        MH_indi[0] = indi_id

                    elif MH_records[-1] == "FAM":
                        fam_id = int(MH_records[1][2:-1])
                #=================================================================================================
                elif MH_records[0] == "1" :
                #=================================================================================================
                    SUBREC_1_KEY = MH_records[1]
                    
                    if SUBREC_1_KEY == "SEX":
                        MH_indi[5] = MH_records[2]
                
                    elif SUBREC_1_KEY == "DEAT":
                        MH_indi[8] = "Y"

                    elif SUBREC_1_KEY == "FAMC":
                        FAMC_row = 'INSERT INTO FAMC (indi_id,fam_id,isAdopted) VALUES (?,?,?);'
                        famc_id = int(MH_records[2][2:-1])
                        FAMC_data = (indi_id,famc_id,"Bio") 
                        sql.execute(FAMC_row, FAMC_data)

                    elif SUBREC_1_KEY == "FAMS":
                        FAMS_row = 'INSERT INTO FAMS (indi_id,fam_id) VALUES (?,?);'
                        FAMS_data = (indi_id,int(MH_records[2][2:-1])) 
                        sql.execute(FAMS_row, FAMS_data)

                    elif SUBREC_1_KEY == "WIFE" or SUBREC_1_KEY == "HUSB" or SUBREC_1_KEY == "CHIL" and fam_id: 
                        FAM_row = f'INSERT INTO {SUBREC_1_KEY} (fam_id,indi_id) VALUES (?,?);'
                        FAM_data = (fam_id,int(MH_records[2][2:-1])) 
                        sql.execute(FAM_row, FAM_data) 

                    elif SUBREC_1_KEY == "OBJE":
                        if isFirst_obje_file : isFirst_obje_file = False
                        else:                                
                            obje_file_row = 'INSERT INTO OBJE (indi_id,form,url,title,date,place,personal) VALUES (?,?,?,?,?,?,?);'
                            obje_file_data = (obje_data[0],obje_data[1],obje_data[2],obje_data[3],obje_data[4],obje_data[5],obje_data[6]) 
                            sql.execute(obje_file_row, obje_file_data)
                        
                        obje_data = [None] * 7
                        obje_data[6] = "N"
                        obje_data[0] = indi_id

                    elif (SUBREC_1_KEY == "EVEN" or 
                          SUBREC_1_KEY == "EDUC" or
                          SUBREC_1_KEY == "OCCU" or
                          SUBREC_1_KEY == "RESI" or
                          SUBREC_1_KEY == "CENS" 
                          ):
                        if isFirst_even : isFirst_even = False
                        else:
                            if (even_data[1] != "EVEN" or 
                                even_data[1] == "EVEN" and even_data[2] in list_events):
                                even_row = 'INSERT INTO INFO (indi_id,even,type,description,date,place,note) VALUES (?,?,?,?,?,?,?);'
                                even_values = (even_data[0],even_data[1],clean_type(even_data[2]),even_data[3],even_data[4],even_data[5],even_data[6]) 
                                sql.execute(even_row, even_values)
                            
                        even_data = [None] * 7
                        even_data[0] = indi_id
                        even_data[1] = SUBREC_1_KEY
                        if len(MH_records) > 2 : 
                            event_description = " ".join(MH_records[2:])
                            even_data[3] = event_description.capitalize()
                        else : event_description = None

                    elif (SUBREC_1_KEY == "NOTE"):
                        
                        if isFirst_bio: isFirst_bio = False
                        else:     
                            bio_texts,bio_hrefs = convert_note(bio_html)
                            for bio_text in bio_texts:
                                bio_row = 'INSERT INTO BIOS (indi_id,note) VALUES (?,?);'
                                bio_data = (bio_indi_id,bio_text) 
                                sql.execute(bio_row, bio_data)
                                
                            for bio_href in bio_hrefs:
                                bio_row = 'INSERT INTO HREF (indi_id,href) VALUES (?,?);'
                                bio_data = (bio_indi_id,bio_href) 
                                sql.execute(bio_row, bio_data)

                        bio_indi_id = indi_id
                        bio_html = " ".join(MH_records[2:])
                        
            #=================================================================================================
                elif MH_records[0] == "2":
            #=================================================================================================
                    SUBREC_2_KEY = MH_records[1]
                    if SUBREC_2_KEY == "SURN":
                        # nom
                        la_data = " ".join(MH_records[2:])
                        la_data = la_data.upper() 
                        la_data = la_data.replace("DE ", "de ") 
                        la_data = la_data.replace("D'", "d'")
                        la_data = la_data.replace("DU ", "du ")
                        la_data = la_data.replace("DE LA ", "de la ")
                        la_data = la_data.replace("de LA ", "de la ")
                        MH_indi[1] = la_data

                    elif SUBREC_2_KEY == "GIVN":
                        # prenom prenoms
                        if len(MH_records) > 2 :  MH_indi[2] = MH_records[2]
                        if len(MH_records) > 3 :  MH_indi[3] = " ".join(MH_records[3:])

                    elif SUBREC_2_KEY == "NICK":
                        # surnom
                        if len(MH_records) > 2 :  MH_indi[4] = " ".join(MH_records[2:])

                    elif SUBREC_1_KEY == "BIRT":

                        if SUBREC_2_KEY == "DATE"    :  MH_indi[6] = clean_date(" ".join(MH_records[2:]))
                        elif SUBREC_2_KEY == "PLAC"  : MH_indi[7]  = clean_place(" ".join(MH_records[2:]))
                            
                    elif SUBREC_1_KEY == "DEAT":  

                        if SUBREC_2_KEY == "DATE"    : MH_indi[9]   = clean_date(" ".join(MH_records[2:]))
                        elif SUBREC_2_KEY == "PLAC"  : MH_indi[10]  = " ".join(MH_records[2:])
                        elif SUBREC_2_KEY == "CAUS":
                            if len(MH_records) > 2   :  MH_indi[11] = " ".join(MH_records[2:])

                    elif SUBREC_2_KEY == "PEDI" and indi_id and famc_id:
                            update_statement = f'UPDATE FAMC SET isAdopted="Adopted" WHERE indi_id = {indi_id} AND fam_id = {famc_id}'
                            sql.execute(update_statement)

                    elif SUBREC_1_KEY == "OBJE":
                        if   SUBREC_2_KEY == "FORM" : obje_data[1] = MH_records[2].lower()
                        elif SUBREC_2_KEY == "FILE" : obje_data[2] = MH_records[2]
                        elif SUBREC_2_KEY == "TITL" : obje_data[3] = MH_records[2]
                        elif SUBREC_2_KEY == "_DATE" : obje_data[4] = MH_records[2]
                        elif SUBREC_2_KEY == "_PLACE" : obje_data[5] = MH_records[2]
                        elif SUBREC_2_KEY == "_PERSONALPHOTO" : obje_data[6] = MH_records[2]

                    elif ((SUBREC_1_KEY     == "EVEN" 
                          or SUBREC_1_KEY   == "EDUC" 
                          or SUBREC_1_KEY   == "RESI" 
                          or SUBREC_1_KEY   == "CENS"
                          or SUBREC_1_KEY   == "OCCU") 
                          ):
                        if    SUBREC_2_KEY == "TYPE" : even_data[2] = " ".join(MH_records[2:])
                        elif  SUBREC_2_KEY == "DATE" : even_data[4] = clean_date(" ".join(MH_records[2:]))
                        elif  SUBREC_2_KEY == "PLAC" : even_data[5] = clean_place(" ".join(MH_records[2:]))
                        elif  SUBREC_2_KEY == "NOTE" : 
                                les_textes,les_href = convert_note(" ".join(MH_records[2:]))
                                if les_textes : even_data[6] = " ".join(les_textes)

                    elif SUBREC_1_KEY      == "NOTE":
                        if SUBREC_2_KEY == "CONC":
                            bio_html = bio_html + " ".join(MH_records[2:])

            # exit loop and final insert

            #INDI
            INDI_row = '''
            INSERT INTO INDI (indi_id,nom,prenom,prenoms,surnom,sexe,bdate,bplace,isdead,ddate,dplace,cause) 
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?);
                        '''
            INDI_data = (MH_indi[0],MH_indi[1],MH_indi[2],MH_indi[3],MH_indi[4],MH_indi[5],
                            MH_indi[6],MH_indi[7],MH_indi[8],MH_indi[9],MH_indi[10],MH_indi[11]) 
            sql.execute(INDI_row, INDI_data)

            #OBJE
            obje_file_row = 'INSERT INTO OBJE (indi_id,form,url,title,date,place,personal) VALUES (?,?,?,?,?,?,?);'
            obje_file_data = (obje_data[0],obje_data[1],obje_data[2],obje_data[3],obje_data[4],obje_data[5],obje_data[6]) 
            sql.execute(obje_file_row, obje_file_data)

            #INFO
            if (even_data[1] != "EVEN" or 
                even_data[1] == "EVEN" and even_data[2] in list_events):
                even_row = 'INSERT INTO INFO (indi_id,even,type,description,date,place,note) VALUES (?,?,?,?,?,?,?);'
                even_values = (even_data[0],even_data[1],clean_type(even_data[2]),even_data[3],even_data[4],even_data[5],even_data[6]) 
                sql.execute(even_row, even_values)
                            
            #BIO

            bio_texts,bio_hrefs = convert_note(bio_html)
            for bio_text in bio_texts:
                bio_row = 'INSERT INTO BIOS (indi_id,note) VALUES (?,?);'
                bio_data = (bio_indi_id,bio_text) 
                sql.execute(bio_row, bio_data)
                
            for bio_href in bio_hrefs:
                bio_row = 'INSERT INTO HREF (indi_id,href) VALUES (?,?);'
                bio_data = (bio_indi_id,bio_href) 
                sql.execute(bio_row, bio_data)

            print(f'{idx} rows added to INDI table')
           
            file_input.close()
            #========================================================================================
            # Close the connection to the database
            connection_obj.commit()
            connection_obj.close()  

    return database_file