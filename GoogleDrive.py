from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from pydrive.files import FileNotUploadedError
from rasa_sdk.interfaces import Tracker
import os

directorio_credenciales = 'credentials_module.json'

# INICIAR SESION
def login():
    GoogleAuth.DEFAULT_SETTINGS['client_config_file'] = directorio_credenciales
    gauth = GoogleAuth()
    gauth.LoadCredentialsFile(directorio_credenciales)
    
    if gauth.credentials is None:
        gauth.LocalWebserverAuth(port_numbers=[8092])
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()
        
    gauth.SaveCredentialsFile(directorio_credenciales)
    credenciales = GoogleDrive(gauth)
    return credenciales

# CREAR ARCHIVO
# def crear_archivo_texto(nombre_archivo,contenido,id_folder):
#     credenciales = login()
#     archivo = credenciales.CreateFile({'title': nombre_archivo,\
#                                        'parents': [{"kind": "drive#fileLink",\
#                                                     "id": id_folder}]})
#     archivo.SetContentString(contenido)
#     archivo.Upload()


# SUBIR UN ARCHIVO A DRIVE
def subir_archivo(ruta_archivo,id_folder):
    credenciales = login()
    archivo = credenciales.CreateFile({'parents': [{"kind": "drive#fileLink",\
                                                    "id": id_folder}]})
    archivo['title'] = ruta_archivo.split("/")[-1]
    archivo['title'] = ruta_archivo.split(os.path.sep)[-1]
    archivo.SetContentFile(ruta_archivo)
    archivo.Upload()