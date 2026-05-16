import os
import sqlite3
import requests # request img from web
import shutil # save img locally
from PIL import Image
# --------------------------------------------------------------
icloud = "/Users/bernardconti/Library/Mobile Documents/com~apple~CloudDocs"
les_WIP = "/Users/bernardconti/LOCAL_TEMP/WIP/"
le_MH_Photos_Bio = icloud+'/MesProgrammes/MH_Photos_Bio/'
#============================================================================================================
class MH_individual(object):
    def __init__(self, dictionary):
        for key, value in dictionary.items():
            setattr(self, key, value)
    def __eq__(self, other):
        if list(self.__dict__.keys()) == list(other.__dict__.keys()): #print("same dict")
            if (list(self.__dict__.values())) == (list(other.__dict__.values())): return True #print("same value")
            else : return False #print("different values")
        else: return False #print("different dict")
#============================================================================================================
MH_none = MH_individual({'indi_id': 0, 'nom': None, 'prenom': None, 'prenoms': None, 'surnom': None, 
        'sexe': None, 'bdate': None, 'bplace': None, 'isdead': None, 'ddate': None, 'dplace': None, 'cause': None})
#============================================================================================================
class photo(object):
    def __init__(self, dictionary):
        for key, value in dictionary.items():
            setattr(self, key, value)
    def __eq__(self, other):
        if list(self.__dict__.keys()) == list(other.__dict__.keys()): #print("same dict")
            if (list(self.__dict__.values())) == (list(other.__dict__.values())): return True #print("same value")
            else : return False #print("different values")
        else: return False #print("different dict")
#============================================================================================================
class MH_couple(object):
    def __init__(self, dictionary):
        for key, value in dictionary.items():
            setattr(self, key, value)
    def __eq__(self, other):
        if list(self.__dict__.keys()) == list(other.__dict__.keys()): #print("same dict")
            if (list(self.__dict__.values())) == (list(other.__dict__.values())): return True #print("same value")
            else : return False #print("different values")
        else: return False #print("different dict")
#============================================================================================================ 
class MH_event(object):
    def __init__(self, dictionary):
        for key, value in dictionary.items():
            setattr(self, key, value)
    def __eq__(self, other):
        if list(self.__dict__.keys()) == list(other.__dict__.keys()): #print("same dict")
            if (list(self.__dict__.values())) == (list(other.__dict__.values())): return True #print("same value")
            else : return False #print("different values")
        else: return False #print("different dict")
#============================================================================================================       
#class MH_individual(object):
#    def __init__(self, **kwargs):
#        self.__dict__.update(kwargs)

#=============================================================================================================
# TEXT
#=============================================================================================================
def text_personne_full(MH_personne,*args):
    isNosurnom = False
    for v in args:
        if v.lower() == "nosurnom" : isNosurnom = True
    if MH_personne != MH_none:
        text = text_personne(MH_personne, "prenom","nom","prenom","surnom" if isNosurnom else "", "bdyear")
    else: text = "MH_none"
    return text
#=============================================================================================================
def text_personne(MH_personne,*args):
#=============================================================================================================
    if len(args) == 0 : 
        print("text_personne sans argument")
        return
    else:
        isAll          = False
        isPrenom      =  False
        isPrenoms     =  False
        isNom         =  False
        isSurnom      =  False
        isSexe        =  False

        isBdate       =  False
        isBdatebold   =  False
        isByear       =  False
        isBdyear      =  False

        isBville      =  False
        isBdepartement=  False
        isBregion     =  False
        isBpays       =  False

        isDdate       =  False
        isDdatebold   =  False
        isDyear       =  False
        isDville      =  False
        isDdepartement=  False
        isDregion     =  False
        isDpays       =  False
        isCause       =  False
        isBDyear       =  False

        #print(args)

        for valeur in args:
            if valeur:
                if isinstance(valeur, str): valeur =valeur.lower()
                if valeur == "all"          : isAll         =  True
                if valeur == "prenom"       : isPrenom      =  True
                if valeur == "prenoms"      : isPrenoms     =  True
                if valeur == "nom"          : isNom         =  True
                if valeur == "surnom"       : isSurnom      =  True
                if valeur == "sexe"         : isSexe        =  True

                if valeur == "bdate"        : isBdate       =  True
                if valeur == "bdatebold"    : isBdatebold   =  True
                if valeur == "byear"        : isByear       =  True
                
                if valeur == "bville"       : isBville      =  True
                if valeur == "bdepartement" : isBdepartement=  True
                if valeur == "bregion"      : isBregion     =  True
                if valeur == "bpays"        : isBpays       =  True

                if valeur == "ddate"        : isDdate       =  True
                if valeur == "ddatebold"    : isDdatebold   =  True
                if valeur == "dyear"        : isDyear       =  True
                if valeur == "dville"       : isDville      =  True
                if valeur == "ddepartement" : isDdepartement=  True
                if valeur == "dregion"      : isDregion     =  True
                if valeur == "dpays"        : isDpays       =  True
                if valeur == "lacause"      : isCause       =  True

                if valeur == "bdyear"       : isBDyear      =  True

        if isAll : return vars(MH_personne)

        le_name=""
        les_names = []
        if MH_personne == MH_none : le_name = "MH_None"
        else:

            if isPrenom         and MH_personne.prenom           : les_names.append(MH_personne.prenom)
            if isNom            and MH_personne.nom              : les_names.append(MH_personne.nom)
            if isPrenoms        and MH_personne.prenoms          : les_names.append(MH_personne.prenoms)
            if isSurnom         and MH_personne.surnom           : les_names.append(MH_personne.surnom)
            if isSexe       and MH_personne.sexe                 : les_names.append(MH_personne.sexe)
            
            # birth dates
            if  MH_personne.bdate:
                if isBdate                                        : les_names.append(MH_personne.bdate)
                if isByear                                        : les_names.append(MH_personne.bdate.split(" ")[-1])
                if isBdatebold                                    : 
                    t = MH_personne.bdate.lower()
                    t = t.replace("janvier","01")
                    t = t.replace("février","02")
                    t = t.replace("mars","03")
                    t = t.replace("avril","04")
                    t = t.replace("mai","05")
                    t = t.replace("juin","06")
                    t = t.replace("juillet","07")
                    t = t.replace("août","08")
                    t = t.replace("septembre","09")
                    t = t.replace("octobre","10")
                    t = t.replace("novembre","11")
                    t = t.replace("décembre","12")
                    t = t.replace(" ","/")
                    temp_date = t.split("/")
                    if len(temp_date) == 3 : les_names.append(f'{int(temp_date[0]):02d}/{temp_date[1]}/<STRONG>{temp_date[-1]}</STRONG>')
                    else : les_names.append(f'<STRONG>{t}</STRONG>')

            #birth places
            if isBville  and MH_personne.bville                           : les_names.append(MH_personne.bville)
            if isBdepartement  and MH_personne.bdepartement               : les_names.append(MH_personne.bdepartement)
            if isBregion  and MH_personne.bregion                         : les_names.append(MH_personne.bregion)
            if isBpays    and MH_personne.bpays                           : les_names.append(MH_personne.bpays)

            # birth dates
            if  MH_personne.ddate:
                if isDdate                                        : les_names.append(MH_personne.ddate)
                if isDyear                                        : les_names.append(f'†{MH_personne.ddate.split(" ")[-1]}')
                if isDdatebold                                    : 
                    t = MH_personne.bdate.lower()
                    t = t.replace("janvier","01")
                    t = t.replace("février","02")
                    t = t.replace("mars","03")
                    t = t.replace("avril","04")
                    t = t.replace("mai","05")
                    t = t.replace("juin","06")
                    t = t.replace("juillet","07")
                    t = t.replace("août","08")
                    t = t.replace("septembre","09")
                    t = t.replace("octobre","10")
                    t = t.replace("novembre","11")
                    t = t.replace("décembre","12")
                    t = t.replace(" ","/")
                    temp_date = t.split("/")
                    if len(temp_date) == 3 : les_names.append(f'{int(temp_date[0]):02d}/{temp_date[1]}/<STRONG>{temp_date[-1]}</STRONG>')
                    else : les_names.append(f'<STRONG>{t}</STRONG>')


            #death places
            if isDville  and MH_personne.dville                           : les_names.append(MH_personne.dville)
            if isDdepartement  and MH_personne.ddepartement               : les_names.append(MH_personne.ddepartement)
            if isDregion  and MH_personne.dregion                         : les_names.append(MH_personne.dregion)
            if isDpays  and MH_personne.dpays                             : les_names.append(MH_personne.dpays)

            if isCause    and MH_personne.cause                           : les_names.append(MH_personne.cause)

            #bdyear
            if isBDyear or isAll:
                if MH_personne.bdate and MH_personne.ddate: 
                    les_names.append(f'({MH_personne.bdate.split(" ")[-1]}-†{MH_personne.ddate.split(" ")[-1]})')
                elif MH_personne.bdate and not MH_personne.ddate: 
                    les_names.append(f'({MH_personne.bdate.split(" ")[-1]})')
                elif not MH_personne.bdate and MH_personne.ddate: 
                    les_names.append(f'(????-†{MH_personne.ddate.split(" ")[-1]})')

            #concatenate les_names
            if les_names: le_name=" ".join(les_names)
        return le_name
#============================================================================================================= 
# PERSONNE
#============================================================================================================= 
def get_MH_indis(sql_obj,le_select):

    # SQL Select 
    MH_indis = []
    #print(le_select)
    sql_obj.execute(le_select)
    for row in sql_obj.fetchall():
        MH_indis.append(MH_individual(dict(row)))

    return MH_indis
#=============================================================================================================
def select_table_join(JOIN_TYPE,TAB_FROM,TAB_TO,common_col):
    SELECT_JOIN = f'{JOIN_TYPE} JOIN {TAB_TO} ON {TAB_FROM}.{common_col}={TAB_TO}.{common_col}'
    return SELECT_JOIN
#=============================================================================================================
def select_table_show_all_rows(sql_obj,table):

    le_select = f'SELECT {table}.* FROM {table}'
    sql_obj.execute(le_select)
    les_rows = []
    for row in sql_obj.fetchall():
        les_rows.append(dict(row))
    return les_rows
#=============================================================================================================
def get_personne_by_indi_id(sql_obj,indi_id):
#=============================================================================================================
    le_select = f"SELECT INDI.* FROM INDI WHERE indi_id = {indi_id}"
    result = get_MH_indis(sql_obj,le_select)

    if result : data = result[0]
    else:       data = MH_none

    return data
#=============================================================================================================
def get_personnes_by_whereclause(sql_obj,la_whereclause):
#=============================================================================================================
    le_select = f"SELECT INDI.* FROM INDI WHERE {la_whereclause}"
    result = get_MH_indis(sql_obj,le_select)
    return result
#=============================================================================================================
def get_personne_data(sql_obj,MH_personne):
#=============================================================================================================
    le_select = f"SELECT INDI.* FROM INDI WHERE {f'indi_id = {MH_personne.indi_id}'}"
    result = get_MH_indis(sql_obj,le_select)

    if result : data = result[0]
    else:       data = MH_none

    return data
#=============================================================================================================
def get_personne_all(sql_obj):
#=============================================================================================================
    le_select = f"SELECT INDI.* FROM INDI"
    result = get_MH_indis(sql_obj,le_select)
    return result
#=============================================================================================================
def get_personne_conjoints(sql_obj,MH_personne):
#-------------------------------------------------------------------------------------------------------------
    le_select =f"""SELECT INDI.* FROM FAMS  
                   {select_table_join("INNER","FAMS","WIFE","fam_id")}
                   {select_table_join("INNER","FAMS","HUSB","fam_id")}
                   INNER JOIN INDI ON (WIFE.indi_id = INDI.indi_id) OR (HUSB.indi_id = INDI.indi_id)
                   WHERE FAMS.indi_id = {MH_personne.indi_id} 
                   and INDI.indi_id != {MH_personne.indi_id}                
                """
    return get_MH_indis(sql_obj,le_select)
#=============================================================================================================
def get_personne_enfants(sql_obj,MH_personne):
#-------------------------------------------------------------------------------------------------------------
    le_select =f"""SELECT INDI.*
                        FROM FAMS 
                        {select_table_join("INNER","FAMS","CHIL","fam_id")}
                        {select_table_join("INNER","CHIL","INDI","indi_id")} 
                        WHERE FAMS.indi_id = {MH_personne.indi_id} 
                        and CHIL.indi_id != {MH_personne.indi_id} 
                        """
    return get_MH_indis(sql_obj,le_select)
#=============================================================================================================
def get_personne_mothers(sql_obj,MH_personne):
#-------------------------------------------------------------------------------------------------------------
    le_select =f"""SELECT DISTINCT INDI.*,FAMC.isAdopted
                    FROM FAMC
                    {select_table_join("INNER","FAMC","WIFE","fam_id")}
                    {select_table_join("INNER","WIFE","INDI","indi_id")}
                    WHERE FAMC.indi_id = {MH_personne.indi_id} 
                    """
    result = get_MH_indis(sql_obj,le_select)

    return result
#=============================================================================================================
def get_personne_fathers(sql_obj,MH_personne):

    le_select =f"""SELECT INDI.*,FAMC.isAdopted
                        FROM FAMC
                        {select_table_join("INNER","FAMC","HUSB","fam_id")}
                        {select_table_join("INNER","HUSB","INDI","indi_id")}
                        WHERE FAMC.indi_id = {MH_personne.indi_id} 
                        """
    result = get_MH_indis(sql_obj,le_select)
    return result
#=============================================================================================================
def get_personne_parents(sql_obj,MH_personne):
#-------------------------------------------------------------------------------------------------------------
    MH_parents = []
    le_select =f"""SELECT FAMC.fam_id,FAMC.isAdopted FROM FAMC 
                WHERE FAMC.indi_id = {MH_personne.indi_id}"""
    
    sql_obj.execute(le_select)
    for row in sql_obj.fetchall():
        fam_id = dict(row)['fam_id']
        isAdopted = dict(row)['isAdopted']

        le_select =f"""SELECT INDI.* FROM HUSB 
                        {select_table_join("INNER","HUSB","INDI","indi_id")}
                        WHERE HUSB.fam_id = {fam_id}"""
        result = get_MH_indis(sql_obj,le_select)
        if len(result) == 1 : h = result[0]
        else : h = MH_none

        le_select =f"""SELECT INDI.* FROM WIFE 
                        {select_table_join("INNER","WIFE","INDI","indi_id")}
                        WHERE WIFE.fam_id = {fam_id}"""
        result = get_MH_indis(sql_obj,le_select)
        if len(result) == 1 : w = result[0]
        else : w = MH_none

        MH_parents.append([h,w,isAdopted])
       
    return MH_parents
#=============================================================================================================
def get_personne_sisbros(sql_obj,MH_personne):
#=============================================================================================================
    le_select_mother_children =f"""SELECT DISTINCT INDI.*
                                    FROM FAMC
                                    {select_table_join("INNER","FAMC","WIFE","fam_id")}
                                    {select_table_join("INNER","WIFE","FAMS","indi_id")}
                                    {select_table_join("INNER","FAMS","CHIL","fam_id")}
                                    {select_table_join("INNER","CHIL","INDI","indi_id")}         
                                    WHERE FAMC.indi_id = {MH_personne.indi_id} and INDI.indi_id != {MH_personne.indi_id} 
                                    """
    le_select_father_children =f"""SELECT DISTINCT INDI.*
                                    FROM FAMC
                                    {select_table_join("INNER","FAMC","HUSB","fam_id")}
                                    {select_table_join("INNER","HUSB","FAMS","indi_id")}
                                    {select_table_join("INNER","FAMS","CHIL","fam_id")}
                                    {select_table_join("INNER","CHIL","INDI","indi_id")}
                                    WHERE FAMC.indi_id = {MH_personne.indi_id} and INDI.indi_id != {MH_personne.indi_id} 
                                    """
    sisbros = get_MH_indis(sql_obj,f'{le_select_mother_children} UNION {le_select_father_children}')

    return sisbros
#=============================================================================================================
def get_personne_oncles(sql_obj,MH_personne,*args):
#------------------------------------------------------------------------------------------------------------- 
    MH_oncles = []
    side = ""

    for valeur in args:
        if valeur:
            if isinstance(valeur, str): valeur =valeur.lower()
            if valeur == "father" : side = "father" 
            if valeur == "mother" : side = "mother"  

    if MH_personne != MH_none : 

        # Branche paternelle
        if side == "father":
            MH_fathers = get_personne_fathers(sql_obj,MH_personne)
            for MH_father in MH_fathers:
                MH_oncles =  get_personne_sisbros(sql_obj,MH_father)

                
        # Branche maternelle
        if side == "mother":
            MH_mothers = get_personne_mothers(sql_obj,MH_personne)
            for MH_mother in MH_mothers: 
                MH_oncles =  get_personne_sisbros(sql_obj,MH_mother)

    else:
    
        print(f"get_personne_oncles:  erreur MH_personne = MH_none")

    return MH_oncles
#=============================================================================================================
def get_personne_events(sql_obj,MH_personne,**kwargs):
    les_events = []

    le_select =f"SELECT INFO.* FROM INFO WHERE INFO.indi_id = {MH_personne.indi_id}"
    for clef, valeur in kwargs.items(): 
        le_select = f"{le_select} AND INFO.{clef} = '{valeur}'"

    sql_obj.execute(le_select)
    for row in sql_obj.fetchall():
        les_events.append(MH_event(dict(row)))
        
    return les_events
#=============================================================================================================
def get_personne_celebrity(sql_obj,MH_personne):
    la_celebrity = None
    les_celebrities = get_personne_events(sql_obj,MH_personne,even ="EVEN",type = "Celebrity")
    if les_celebrities : la_celebrity = les_celebrities[0].description
    return la_celebrity
#=============================================================================================================
def get_personne_bios(sql_obj,MH_personne):
    les_bios = []

    le_select =f"SELECT BIOS.* FROM BIOS WHERE BIOS.indi_id = {MH_personne.indi_id}"
    sql_obj.execute(le_select)
    for row in sql_obj.fetchall():
        les_bios.append(row["note"])
        
    return les_bios
#=============================================================================================================
def get_personne_hrefs(sql_obj,MH_personne):
    les_hrefs = []

    le_select =f"SELECT HREF.* FROM HREF WHERE HREF.indi_id = {MH_personne.indi_id}"
    sql_obj.execute(le_select)
    for row in sql_obj.fetchall():
        les_hrefs.append(row["href"])
        
    return les_hrefs
#=============================================================================================================
def get_personne_descendants(sql_obj,MH_personne,MH_couples_descendant,n_level,max_level):
#------------------------------------------------------------------------------------------------------------- 
# MH_get_personne_descendants = [n_level,MH_personne,MH_personne_conjoint,Status filiation = Biologique ou Alliance])
    n_level = n_level +1
    if n_level <= max_level :

        MH_conjoints = get_personne_conjoints(sql_obj,MH_personne)
        if MH_conjoints : 
    #=============================================================================================================        
    # # on mémorise tous les enfants biologiques de la personne
            # on recherche les conjoints
            for MH_conjoint in MH_conjoints:

                # on enregistre la personne avec son conjoint
                MH_couples_descendant.append( MH_couple({"level" :n_level, "adult1": MH_personne, "adult2" : MH_conjoint}) )

                #on prepare la boucle suivante
                MH_enfants_conjoint = get_personne_enfants(sql_obj,MH_conjoint)

                for MH_enfant_conjoint in MH_enfants_conjoint:
    
                    # lancemement de l'itération
                    get_personne_descendants(sql_obj,MH_enfant_conjoint,MH_couples_descendant,n_level,max_level)
        else:
#---------- on enregistre la personne sans conjoint
            MH_couples_descendant.append( MH_couple({"level" :n_level, "adult1": MH_personne, "adult2" : MH_none}) )
#------------------------------------------------------------------------------------------------------------- 
    return MH_couples_descendant
#=============================================================================================================
def get_personne_descendants_n_level(sql_obj,MH_personne,n):
#-------------------------------------------------------------------------------------------------------------
    MH_descendants = get_personne_descendants(sql_obj,MH_personne,[],0,n)

    n_MH_adults = 0       
    for MH_descendant in MH_descendants:
        n_MH_adults = n_MH_adults+1
        if MH_descendant["adult2"] != MH_none : n_MH_adults = n_MH_adults + 1

    if n_MH_adults < 60: 
        return MH_descendants,n
    else: 
        MH_descendants,n = get_personne_descendants_n_level(MH_personne,n-1)
#-------------------------------------------------------------------------------------------------------------
        return MH_descendants,n
#=============================================================================================================
def get_personne_ascendants(sql_obj,MH_personne,MH_ascendants,n_level,n_level_max):
#-------------------------------------------------------------------------------------------------------------
    if MH_personne != MH_none: 
        n_level = n_level + 1
        if n_level == 1:  
            MH_ascendants.append(MH_couple({"level":n_level, "adult1":MH_personne , "adult2":MH_none, "bio":"bio" }))
            MH_ascendants = get_personne_ascendants(sql_obj,MH_personne,MH_ascendants,n_level,n_level_max)  
            
        elif n_level <= n_level_max: 
            for MH_parent in get_personne_parents(sql_obj,MH_personne):
                adult1 = MH_parent[0]
                adult2 = MH_parent[1]
                isAdopted = MH_parent[2]
                MH_ascendants.append(MH_couple({"level":n_level, "adult1":adult1 , "adult2":adult2, "bio":isAdopted }))
                MH_ascendants = get_personne_ascendants(sql_obj,adult1,MH_ascendants,n_level,n_level_max)

                MH_ascendants.append(MH_couple({"level":n_level, "adult1":adult2 , "adult2":adult1, "bio":isAdopted }))
                MH_ascendants = get_personne_ascendants(sql_obj,adult2,MH_ascendants,n_level,n_level_max)
    #else:
    #    print("MH_personne = None")

    return MH_ascendants
#================================================================
def get_personne_ascendants_n_level(sql_obj,MH_personne,n_line_max,n_level):
#-------------------------------------------------------------------------------------------------------------    
    MH_ascendants = get_personne_ascendants(sql_obj,MH_personne,[],0,n_level)
    if len(MH_ascendants) - 1 <= n_line_max: return MH_ascendants,n_level
    else: 
        MH_ascendants,n_level = get_personne_ascendants_n_level(MH_personne,n_line_max,n_level-1)
        return MH_ascendants,n_level
#=============================================================================================================
def get_personne_photo(sql_obj,MH_personne,*args):
#-------------------------------------------------------------------------------------------------------------
    MH_photos = []
    if MH_personne != MH_none:
        add_select = 'AND OBJE.personal == "N"'
        for value in args:
            if value == "photo_id" : add_select = 'AND OBJE.personal = "Y"'

        if MH_personne != MH_none:
            le_select =f"""SELECT OBJE.url,OBJE.title,OBJE.date,OBJE.place
                                    FROM OBJE
                                    WHERE OBJE.indi_id = {MH_personne.indi_id}
                                    AND OBJE.form = "jpg"
                                    {add_select}"""
            sql_obj.execute(le_select)
            for row in sql_obj.fetchall():
                MH_photos.append(photo(dict(row)))
    return MH_photos
#=============================================================================================================
def get_personne_photoID_file(sql_obj,MH_personne):
#-------------------------------------------------------------------------------------------------------------       
    images_url = get_personne_photo(sql_obj,MH_personne,"photo_id")
    if len(images_url) > 0 : 

        if len(images_url) > 1 : print("get_personne_photoID_file: étrange, plusieurs photo_ids, on ne garde que la première")

        image_url = images_url[0].url
        image_file = image_url.split('/')[-1]
        img_path =  icloud+'/MesProgrammes/MH_Photos/'+image_file
#------ test if file allready loaded on disk, if not load from URL Link
        if not os.path.isfile(img_path):
            res = requests.get(image_url, stream = True)
            if res.status_code == 200:
                with open(img_path,'wb') as f:
                    shutil.copyfileobj(res.raw, f)
                    print ("--> download successfull"+ img_path)
    else:
        if MH_personne.sexe == "M": img_path =  icloud+'/MesProgrammes/MH_Photos/personne_M.png'
        else: img_path =  icloud+'/MesProgrammes/MH_Photos/personne_F.png'
    
    return img_path
#=============================================================================================================
def get_personne_MHphotos(sql_obj,MH_personne,*args):
#-------------------------------------------------------------------------------------------------------------
    buffer_photos = []
    if MH_personne != MH_none:
        #-------------------------------------------------------------------------------------------------------------
        mode = "all"
        for valeur in args:
            if valeur:
                if isinstance(valeur, str): valeur =valeur.lower()
                if valeur == "bio" : mode = valeur  
        #-------------------------------------------------------------------------------------------------------------
        buffer_photos = []
        les_photos = get_personne_photo(sql_obj,MH_personne)

    #-------------------------------------------------------------------------------------------------------------
        if les_photos:
            for la_photo in les_photos:
    #-------------------------------------------------------------------------------------------------------------              
                photo_url = la_photo.url
                photo_title = la_photo.title
                isBio = False
                if mode == "bio":
                    if photo_title:
                        if "BIO-" in photo_title :
                            n = photo_title.split("_")[0].replace("BIO-","")
                            if n: 
                                if n.isnumeric(): 
                                    photo_title =  f'{int(n):02d}'
                                    isBio = True
    #--------------------------------------------------------------------------------------------------
                if (photo_url and ((isBio and mode == "bio") or mode =="all")):

                    le_watermark = []
    # date
                    photo_date = la_photo.date  
                    if not photo_date: photo_date = "9999"
                    else: photo_date = photo_date.split(" ")[-1]
                    le_watermark.append([photo_date,"date"])
    # title 
                    le_watermark.append([photo_title,"titre"])
    # place
                    photo_lieu = la_photo.place
                    if photo_lieu : le_watermark.append([photo_lieu,"place"])

                    le_dir = f'{le_MH_Photos_Bio}INDI_{MH_personne.indi_id:05d}/'
                    if not os.path.isdir(le_dir):
                        os.makedirs(le_dir)
                        print("création du répertoire: ",le_dir)
    #-----  Image URL 
                    if photo_url:
                        image_file = photo_url.split('/')[-1]
                        img_path =  le_dir+image_file
    #------ test if file allready loaded on disk, if not load from URL Link
                        if not os.path.isfile(img_path):
                            res = requests.get(photo_url, stream = True)
                            if res.status_code == 200:
                                with open(img_path,'wb') as f:
                                    shutil.copyfileobj(res.raw, f)
                                    print ("--> download successfull"+ img_path)
                            else : 
                                print("erreur extration du fichier url",photo_url)
                                return buffer_photos
    #-- initiage image from local file
                    img = Image.open(img_path)
    #-------------------------------------------------------------------------------------------------------------
                    buffer_photos.append([f'{photo_title}"|"{photo_date}',img_path,img.width,img.height,le_watermark])                       
    return buffer_photos
#=============================================================================================================
def get_personne_entourage(sql_obj,MH_personne,n_level_max,loption): 
#-------------------------------------------------------------------------------------------------------------
    MH_couples = []
    MH_entourages_new = []
    le_level_max = 0
    #---------------------------------------------------------------------------------------------------------
    if loption == "ascendant" : 
        MH_couples = get_personne_ascendants(sql_obj,MH_personne,[],0,n_level_max)
        if len(MH_couples) == 1: MH_couples = []
    #---------------------------------------------------------------------------------------------------------
    elif loption ==  "fratrie":
            MH_sisbros = get_personne_sisbros(sql_obj,MH_personne)
            if not MH_sisbros : return MH_entourages_new,le_level_max
            for item in MH_sisbros:
                MH_couples = MH_couples + get_personne_descendants(sql_obj,item,[],0,n_level_max)
    #---------------------------------------------------------------------------------------------------------
    elif loption == "couzpater":
        MH_oncles = get_personne_oncles(sql_obj,MH_personne,"father")
        if not MH_oncles : return MH_entourages_new,le_level_max
        for item in MH_oncles:
            MH_couples = MH_couples + get_personne_descendants(sql_obj,item,[],0,n_level_max)
    #---------------------------------------------------------------------------------------------------------
    elif loption == "couzmater":
        MH_oncles = get_personne_oncles(sql_obj,MH_personne,"mother")
        if not MH_oncles : return MH_entourages_new,le_level_max
        for item in MH_oncles:
            MH_couples = MH_couples + get_personne_descendants(sql_obj,item,[],0,n_level_max)
    #---------------------------------------------------------------------------------------------------------
    elif loption == "descendant":     
        MH_couples = get_personne_descendants(sql_obj,MH_personne,[],0,n_level_max)
    #---------------------------------------------------------------------------------------------------------
    else :
        print("get_personne_entourage",loption)

    #---------------------------------------------------------------------------------------------------------
    if MH_couples : 
        if loption == "ascendant":
            for MH_couple in MH_couples:
                    #print(MH_couple)
                    le_level_max = max(MH_couple.level,le_level_max)
                    la_bdate = MH_couple.adult1.bdate
                    if la_bdate : la_year = MH_couple.adult1.bdate.split(" ")[-1]
                    else : la_year ="????"
                    #les_parents = get_personne_parents(sql_obj,MH_couple.adult1)
                    MH_entourages_new.append([MH_couple.level,MH_couple.adult1,MH_couple.adult2,la_year,MH_couple.adult1.sexe,MH_couple.bio])

        else:
            if MH_couples[0].adult2 != MH_none or loption != "descendant":
                for MH_couple in MH_couples:
                    le_level_max = max(MH_couple.level,le_level_max)
                    if MH_couple.adult1.bdate :
                        la_year = MH_couple.adult1.bdate.split(" ")[-1]
                    else: la_year ="????"
                    MH_entourages_new.append([MH_couple.level,MH_couple.adult1,MH_couple.adult2,la_year])
            #else:
                #MH_couple = []
        
    return MH_entourages_new,le_level_max
#=============================================================================================================
def get_personne_lignées(sql_obj,MH_personne):
#-------------------------------------------------------------------------------------------------------------

    if MH_personne != MH_none:
        n_level_max = 0
        lignées = []
        MH_ascendants = get_personne_ascendants(sql_obj,MH_personne,[],0,99)
        #MH_ascendants = MH_couple({"level":n_level, "adult1":adult1 , "adult2":adult2, "bio":isAdopted }))

        for MH_couple in MH_ascendants : 
            n_level_max = max(n_level_max,MH_couple.level)
        lignée = [MH_none]* n_level_max
        lignées_previous = lignée
        for MH_couple in MH_ascendants : 

            n_level     = MH_couple.level
            MH_adult1   = MH_couple.adult1

            if MH_adult1.sexe == "F" : 
                lignées_previous = lignée
                lignées.append(lignée)
                lignée = [MH_none]* n_level_max
                for i in range(0,n_level-1):
                    lignée[i]=lignées_previous[i]

            lignée[n_level-1] = MH_adult1
            
        lignées.append(lignée)
#-------------------------------------------------------------------------------------------------------------
    return lignées
#=============================================================================================================
def get_personne_lignée(MH_personne,lignées):
#-------------------------------------------------------------------------------------------------------------
    lignée = []
    idx = 0
    if MH_personne != MH_none and lignées:
        for line in lignées:
            lignée = []
            for idx,MH_item in enumerate(line):
                if MH_item == MH_none: break
                lignée.append(MH_item)
                if MH_item == MH_personne: return lignée,idx
#-------------------------------------------------------------------------------------------------------------
    return lignée,idx
#=============================================================================================================
def is_personne_in_MH_list(MH_personne,MH_list):
    if MH_personne != MH_none :
        for MH_item in MH_list:
            if MH_item == MH_personne : return True
    return  False
# COUPLE
#=============================================================================================================
def get_couple_enfants(sql_obj,MH_adult1,MH_adult2):
#=============================================================================================================
    #les enfants BIO du couples
    le_select1 =f"""SELECT  INDI.*
                    FROM FAMS a
                    JOIN FAMS b ON a.fam_id = b.fam_id
                    {select_table_join("INNER","b","CHIL","fam_id")}
                    {select_table_join("INNER","CHIL","INDI","indi_id")} 
                    WHERE a.indi_id = {MH_adult1.indi_id} AND b.indi_id = {MH_adult2.indi_id}   
                    """
    #les enfants bio du Adult2
    le_select2 =f"""SELECT  INDI.*
                    FROM FAMS 
                    {select_table_join("INNER","FAMS","CHIL","fam_id")}
                    {select_table_join("INNER","CHIL","INDI","indi_id")} 
                    WHERE FAMS.indi_id = {MH_adult2.indi_id}   
                    """
        
    #union des 2 resultats
    return get_MH_indis(sql_obj,f'{le_select1} UNION {le_select2}')



    MH_enfants = []
    n_enfant = 0
    for item in get_couple_descendants([],MH_adult1,MH_adult2,0,2,False):
        if item[0] == 2 : 
            n_enfant = n_enfant+1
            MH_enfants.append(item)
    return MH_enfants,n_enfant
#=============================================================================================================
# COUPLE
#=============================================================================================================
def get_couple_descendants_n_level(MH_adult1,MH_adult2,n):
#-------------------------------------------------------------------------------------------------------------
    MH_descendants = get_couple_descendants([],MH_adult1,MH_adult2,0,n,False)

    n_MH_adults = 0       
    for MH_descendant in MH_descendants:
        n_MH_adults = n_MH_adults+1
        if MH_descendant[2] != "???" : n_MH_adults = n_MH_adults + 1

    if n_MH_adults < 60: 
        return MH_descendants,n
    else: 
        MH_descendants,n = get_couple_descendants_n_level(MH_adult1,MH_adult2,n-1)
#-------------------------------------------------------------------------------------------------------------
        return MH_descendants,n
#============================================================================================================= 
def get_couple_descendants(MH_get_personne_descendants,MH_adult1,MH_adult2,n_level,max_level,isBIO):
#------------------------------------------------------------------------------------------------------------- 
# MH_get_personne_descendants = [n_level,MH_personne,MH_personne_conjoint,Status filiation = Biologique ou Alliance])
    if MH_adult2 != "???":
        n_level = n_level +1
        if n_level <= max_level :
    #-------------- on enregistre la personne avec son conjoint
            MH_get_personne_descendants.append([n_level,MH_adult1,MH_adult2,isBIO])

    #-------------- on recherche les enfants biologiques du conjoint
            MH_conjoint_enfants = MH_adult2.sub_tags('FAMS/CHIL')

    #---------- on boucle sur les enfants commun entre la personne et le conjoint uniquement
            if MH_conjoint_enfants :
                for MH_conjoint_enfant in MH_conjoint_enfants:

    #---------------------- on recherche le parent2 de l'enfant du conjoint (parent1)
                    
                    if MH_adult2.sex == "M" : MH_conjoint_enfant_parents_2 = MH_conjoint_enfant.sub_tags('FAMC/WIFE')
                    else:MH_conjoint_enfant_parents_2 = MH_conjoint_enfant.sub_tags('FAMC/HUSB')

    #---------------------- on boucle si  la personne est dans la liste des MH_conjoint_enfant_parents_2
                    if in_MH_list(MH_adult1,MH_conjoint_enfant_parents_2):
    #-------------------------- on regarde sil'enfant du conjoint est dans la liste des enfants BIO
                        #if in_MH_list(MH_conjoint_enfant,MH_personne_enfants):isBIO = True
                        #else: isBIO = False

                        get_personne_descendants(MH_get_personne_descendants,MH_conjoint_enfant,n_level,max_level,isBIO)
            else:
    #---------- on enregistre la personne sans conjoint
                MH_get_personne_descendants.append([n_level,MH_adult1,"???",isBIO])
#------------------------------------------------------------------------------------------------------------- 
    return MH_get_personne_descendants
#=============================================================================================================


