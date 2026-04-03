import os
import sqlite3
import requests # request img from web
import shutil # save img locally
from PIL import Image
# --------------------------------------------------------------
icloud = "/Users/bernardconti/Library/Mobile Documents/com~apple~CloudDocs"
les_WIP = "/Users/bernardconti/LOCAL_TEMP/WIP/"
le_MH_Photos_Bio = icloud+'/MesProgrammes/MH_Photos_Bio/'
#--------------------------------------------------------------
#Gray= "#EFEEEF"
#GrayRow= "#EFEEEF"
#Black ="#000000"
#White = "#FFFFFF"
#couleur_homme ="#ffcc99"
#couleur_femme="#ccccff"
#font_texte ="arial narrow"
#couleur_cible = "#30DD30"
#Gray1= "#b7b1b1"#F2F2F2
#GraySide = "#c5c5c5"
#Gray3 ="#D9D6D6"
#Gray4 = "#E1DEDEEF"
#couleur_chemin = "#EC5800"
#couleur_titre ="#EB742F"
#Green = "#B4E1C0"
#Gray2= "#fde9d9"
#Silver = "#A9A6A6" 
#cell11 = "#F8CEB8"
#cell12 = "#D9D9D9"
#cell21 = "#FDE8DD"
#cell22 = "#F2F2F2"
#Bleu = "#375e94"
#Bleu = "#272170"
#Darkblue = "#166082"
#Gris_clair = "#EFEFEF"
#ligne = "#375e94"
#descendant = "#000000"
#red = "#FF0000"
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
MH_none = MH_individual({'indi_id': 0, 'nom': None, 'prenom': None, 'prenoms': None, 'surnom': None, 
        'sexe': None, 'bdate': None, 'bplace': None, 'isdead': None, 'ddate': None, 'dplace': None, 'cause': None})
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
        text = f"{MH_personne.prenom if MH_personne.prenom else "prenom_???"} {MH_personne.nom if MH_personne.nom else "nom_???"}{" , "+MH_personne.prenoms if MH_personne.prenoms else ""}{' "'+MH_personne.surnom+'"' if MH_personne.surnom and not isNosurnom else ''}"
    else: text = "MH_none"
    return text
#=============================================================================================================
def text_personne(MH_personne,*args):
#=============================================================================================================
    le_name=""
    les_names = []
    isOption = [False]*17

    for valeur in args:
        if valeur:
            if isinstance(valeur, str): valeur =valeur.lower()
            if valeur == "prenom"       : isOption[0]   =  True
            if valeur == "prenoms"      : isOption[1]   =  True
            if valeur == "nom"          : isOption[2]   =  True
            if valeur == "surnom"       : isOption[3]   =  True
            if valeur == "sexe"         : isOption[4]   =  True
            if valeur == "bdate"        : isOption[5]   =  True
            if valeur == "byear"        : isOption[6]   =  True
            if valeur == "bplace"       : isOption[7]   =  True
            if valeur == "bcity"        : isOption[8]   =  True
            if valeur == "ddate"        : isOption[9]   =  True
            if valeur == "dyear"        : isOption[10]  =  True
            if valeur == "dplace"       : isOption[11]  =  True
            if valeur == "dcity"        : isOption[12]  =  True
            if valeur == "lacause"      : isOption[13]  =  True
            if valeur == "bdyear"       : isOption[14]  =  True
            if valeur == "bcountry"     : isOption[15]  =  True
            if valeur == "bregion"      : isOption[16]  =  True

    le_name = ""
    if MH_personne == MH_none : le_name = "MH_None"
    else:
        byear = None
        dyear = None
        if MH_personne.bdate : la_bdate = MH_personne.bdate
        else : la_bdate = "?? ??? ????"
        byear = la_bdate.split(" ")[-1]

        if MH_personne.isdead:
            if MH_personne.ddate : la_ddate = MH_personne.ddate
            else: la_ddate = "?? ??? ????"
            dyear = "†"+la_ddate.split(" ")[-1]
        else: la_ddate = None

        bcity = None
        bregion = None
        bcountry = None

        if MH_personne.bplace : 
            temp_place = MH_personne.bplace.split(",")
            bcity = temp_place[0].lstrip().title()
            if len(temp_place) > 1 : bregion = temp_place[1].lstrip().title()
            if len(temp_place) > 2 : bcountry = temp_place[-1].lstrip().upper()

        if MH_personne.dplace : dcity = MH_personne.dplace.split(",")[0].lstrip().capitalize()

        if isOption[0] and MH_personne.prenom       : les_names.append(MH_personne.prenom)
        if isOption[1] and MH_personne.prenoms      : les_names.append(MH_personne.prenoms)
        if isOption[2] and MH_personne.nom          : les_names.append(MH_personne.nom)
        if isOption[3] and MH_personne.surnom       : les_names.append(f'"{MH_personne.surnom}"')
        if isOption[4] and MH_personne.sexe         : les_names.append(MH_personne.sexe)
        if isOption[5] and MH_personne.bdate        : les_names.append(la_bdate)
        if isOption[6] and byear                    : les_names.append(byear)
        if isOption[7] and MH_personne.bplace       : les_names.append(MH_personne.bplace)
        if isOption[8] and MH_personne.bplace       : les_names.append(bcity)
        if isOption[9] and MH_personne.ddate        : les_names.append(la_ddate)
        if isOption[10] and dyear                   : les_names.append(dyear)
        if isOption[11] and MH_personne.dplace      : les_names.append(MH_personne.dplace)
        if isOption[12] and MH_personne.dplace      : les_names.append(dcity)
        if isOption[13] and MH_personne.cause       : les_names.append(MH_personne.cause)
        if isOption[14] :
            if byear and dyear : les_names.append(f'({byear}-{dyear})')
            elif byear and not dyear : les_names.append(f'({byear})')

        
        if isOption[15] and MH_personne.bplace: les_names.append(f'{bcountry}')
        if isOption[16] and MH_personne.bplace: les_names.append(f'{bregion}')
        
    

        if les_names: le_name=" ".join(les_names)
    return le_name
#=============================================================================================================
def search_country(bplace): 
#-------------------------------------------------------------------------------------------------------------
    bcountry = "FRANCE"
# liste des pays
    from pays import Countries
    countries = Countries('fra')
    les_pays = []
    for country in countries:  # générateur
        les_pays.append(country.name.lower())

    if bplace : 
        country = "FRANCE"
        for le_pays in les_pays:
            if le_pays in bplace.lower():
                bcountry = le_pays.upper()
                break

    return bcountry
#=============================================================================================================
def text_couple(MH_adult1,MH_adult2,*args):
#=============================================================================================================
    le_mode = "complet"
    for c in args:
        if c =="simple" : le_mode = "simple"

    le_texte = ""
    temp_le_texte = []
    if MH_adult1 == MH_none: le_texte = text_personne_full(MH_adult1)
    else:        
        #=====================================================================================================
        if le_mode == "simple":
            if MH_adult1.prenom       : temp_le_texte.append(f'<strong>{MH_adult1.prenom}</strong>')
            if MH_adult1.nom          : temp_le_texte.append(f'{MH_adult1.nom}')
            if MH_adult1.bdate        : temp_le_texte.append(MH_adult1.bdate.split(" ")[-1])
            if MH_adult1.bplace       : temp_le_texte.append(f'à {MH_adult1.bplace.split(",")[0].capitalize()}')

            if MH_adult2 != MH_none:
                temp_le_texte.append(f'et')
                if MH_adult2.prenom       : temp_le_texte.append(f'<strong>{MH_adult2.prenom}</strong>')
                if MH_adult2.nom          : temp_le_texte.append(f'{MH_adult2.nom}')
                if MH_adult2.bdate        : temp_le_texte.append(MH_adult2.bdate.split(" ")[-1])
                if MH_adult2.bplace       : temp_le_texte.append(f'à {MH_adult2.bplace.split(",")[0].capitalize()}')
        #======================================================================================================
        # mode complet
        else:
            conjugaison = "e" if MH_adult1.sexe == "F" else ""
            if MH_adult1.prenom       : temp_le_texte.append(f'<strong>{MH_adult1.prenom}</strong>')
            if MH_adult1.nom          : temp_le_texte.append(f'{MH_adult1.nom}')
            if MH_adult1.prenoms       : temp_le_texte.append(f'{MH_adult1.prenoms}')
            if MH_adult1.surnom       : temp_le_texte.append(f'"{MH_adult1.surnom}"')
            if MH_adult1.bdate        : temp_le_texte.append(f',né{conjugaison} le {MH_adult1.bdate}')
            if MH_adult1.bplace       : temp_le_texte.append(f'à {MH_adult1.bplace.split(",")[0].capitalize()}')
            if MH_adult1.ddate        : temp_le_texte.append(f'†{MH_adult1.bdate}')
            if MH_adult1.dplace       : temp_le_texte.append(f'à {MH_adult1.dplace.split(",")[0].capitalize()}')

            if MH_adult2 != MH_none:
                conjugaison = "e" if MH_adult2.sexe == "F" else ""
                temp_le_texte.append(f'et')
                if MH_adult2.prenom       : temp_le_texte.append(f'<strong>{MH_adult2.prenom}</strong>')
                if MH_adult2.nom          : temp_le_texte.append(f'{MH_adult2.nom}')
                if MH_adult2.prenoms       : temp_le_texte.append(f'{MH_adult2.prenoms}')
                if MH_adult2.surnom       : temp_le_texte.append(f'"{MH_adult2.surnom}"')
                if MH_adult2.bdate        : temp_le_texte.append(f',né{conjugaison} le {MH_adult2.bdate}')
                if MH_adult2.bplace       : temp_le_texte.append(f'à {MH_adult2.bplace.split(" ")[0].capitalize()}')
                if MH_adult2.ddate        : temp_le_texte.append(f'†{MH_adult2.bdate}')
                if MH_adult2.dplace       : temp_le_texte.append(f'à {MH_adult2.dplace.split(" ")[0].capitalize()}')

        if temp_le_texte: le_texte=" ".join(temp_le_texte)
    return le_texte 
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
    idx =0
    le_select = f'SELECT {table}.* FROM {table}'
    sql_obj.execute(le_select)
    for idx,row in enumerate(sql_obj.fetchall()):
        print(dict(row))
    return idx+1
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
            if MH_couples[0].adult2 != MH_none:
                for MH_couple in MH_couples:
                    #MH_couple)
                    le_level_max = max(MH_couple.level,le_level_max)
                    if MH_couple.adult1.bdate :
                        la_year = MH_couple.adult1.bdate.split(" ")[-1]
                    else: la_year ="????"
                    #les_parents = get_personne_parents(sql_obj,MH_couple.adult1)
                    MH_entourages_new.append([MH_couple.level,MH_couple.adult1,MH_couple.adult2,la_year])
            else:
                MH_couple = []
        
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


#OLD
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

