import arcpy

def listaDrog(lista_drog):
    with arcpy.da.SearchCursor('drogi_wodne', ['MS_KOD']) as cur_drogi0:
        for row_drogi in cur_drogi0:
            lista_drog.append(row_drogi[0])
    del cur_drogi0


def listaElek(lista_elek):
    with arcpy.da.SearchCursor('elektrownie_opis', ['MS_KOD']) as cur_elek0:
        for row_elek in cur_elek0:
            lista_elek.append(row_elek[0])
    del cur_elek0

def listaPob(lista_pob):
    with arcpy.da.SearchCursor('pobory', ['Kod_JCWP']) as cur_pob0:
        for row_pob in cur_pob0:
            lista_pob.append(row_pob[0])
    del cur_pob0


def listaOch(lista_och):
    with arcpy.da.SearchCursor('Obsz_min_pow', ['MS_KOD']) as cur_och0:
        for row_och in cur_och0:
            lista_och.append(row_och[0])
    del cur_och0

def listaLch(lista_lch):
    with arcpy.da.SearchCursor('Ludnosc_chroniona', ['MS_KOD']) as cur_lch0:
        for row_lch in cur_lch0:
            lista_lch.append(row_lch[0])
    del cur_lch0


def listaOc(lista_oc):
    with arcpy.da.SearchCursor('ob_cenne', ['MS_KOD']) as cur_oc0:
        for row_oc in cur_oc0:
            lista_oc.append(row_oc[0])
    del cur_oc0

def listaZoo(lista_zoo):
    with arcpy.da.SearchCursor('zoo', ['MS_KOD']) as cur_zoo0:
        for row_zoo in cur_zoo0:
            lista_zoo.append(row_zoo[0])
    del cur_zoo0

def listaCm(lista_cm):
    with arcpy.da.SearchCursor('cmentarze', ['MS_KOD']) as cur_cm0:
        for row_cm in cur_cm0:
            lista_cm.append(row_cm[0])
    del cur_cm0

def listaKap(lista_kap):
    with arcpy.da.SearchCursor('kapieliska', ['MS_KOD']) as cur_kap0:
        for row_kap in cur_kap0:
            lista_kap.append(row_kap[0])
    del cur_kap0

def listaOcz(lista_oczysz):
    with arcpy.da.SearchCursor('oczyszczalnie', ['MS_KOD']) as cur_oczysz0:
        for row_oczysz in cur_oczysz0:
            lista_oczysz.append(row_oczysz[0])
    del cur_oczysz0

def listaUj(lista_uj):
    with arcpy.da.SearchCursor('ujecia_wod', ['MS_KOD']) as cur_uj0:
        for row_uj in cur_uj0:
            lista_uj.append(row_uj[0])
    del cur_uj0

def listaSkl(lista_sklad):
    with arcpy.da.SearchCursor('skladowiska_odpadow', ['MS_KOD']) as cur_sklad0:
        for row_sklad in cur_sklad0:
            lista_sklad.append(row_sklad[0])
    del cur_sklad0

def listaZak(lista_zaklad):
    with arcpy.da.SearchCursor('zaklady_przem', ['MS_KOD']) as cur_zaklad0:
        for row_zaklad in cur_zaklad0:
            lista_zaklad.append(row_zaklad[0])
    del cur_zaklad0


def listaMel(lista_mel):
    with arcpy.da.SearchCursor('ob_mel', ['MS_KOD']) as cur_mel0:
        for row_mel in cur_mel0:
            lista_mel.append(row_mel[0])
    del cur_mel0

def listaKru(lista_krusz):
    with arcpy.da.SearchCursor('kruszywa_pob', ['MS_KOD']) as cur_krusz0:
        for row_krusz in cur_krusz0:
            lista_krusz.append(row_krusz[0])
    del cur_krusz0

def listaPrz(lista_prze):
    with arcpy.da.SearchCursor('przerzuty_wod', ['MS_KOD']) as cur_prze0:
        for row_prze in cur_prze0:
            lista_prze.append(row_prze[0])
    del cur_prze0

def listaGor(lista_gor):
    with arcpy.da.SearchCursor('gornictwo', ['MS_KOD']) as cur_gor0:
        for row_gor in cur_gor0:
            lista_gor.append(row_gor[0])
    del cur_gor0