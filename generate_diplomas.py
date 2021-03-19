import gspread, configparser, os, datetime
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
from mail_sender import send_email 

def log(l):
    """ Log d'exécution """
    with open(f'logs/log_{d}.log', 'a') as wf:
        wf.write(l)

def generate_report(success, fails):
    """ Génère le rapport d'exécution """

    log(f"Rapport d'exécution - {d}\n\n{len(sheet_list)} lignes au total sur le GSheets\n{len(name_list)} lignes marquées comme à traiter\n{len(success)} mails envoyés avec succès\n{len(fails)} échecs lors de l'envoi\n\n")

    if len(success) > 0:
        tw = "Nom, prénom et email des personnes à qui les attestations ont été correctement envoyées:\n"
        for f in success:
            tw += f'{f[0]} {f[1]} - {f[2]}\n'
        tw += '\n'
        log(tw)

    if len(fails) > 0:
        tw = "Nom, prénom et email des personnes à qui les attestations n'ont pas pu être envoyées:\n"
        for f in fails:
            tw += f'{f[0]} {f[1]} - {f[2]}\n'
        tw += '\n'
        log(tw)


def get_text_dimensions(text_string, font):
    """ Retourne les dimensions du texte avec la police et taille donnée """
    ascent, descent = font.getmetrics()
    text_width = font.getmask(text_string).getbbox()[2]
    text_height = font.getmask(text_string).getbbox()[3] + descent

    return (text_width, text_height)


def generate_image(nom, prenom):
    """ Enregistre l'attestation avec le nom et prénom inscrit """ 
    img = Image.open(diploma_template)
    draw = ImageDraw.Draw(img)
    nom = nom.upper()
    prenom = prenom.upper()

    txt = f'{prenom} {nom}'
    font = ImageFont.truetype("NotoSerifCJK-Bold.ttc", 140)

    # Calcul de la position du texte
    w_max = 3508
    txt_width = get_text_dimensions(txt, font)[0]
    perfect = (w_max - txt_width) / 2 # - 20

    draw.text((perfect, 1500), txt, (255, 255, 255), font=font)

    path = f'diplomas/{prenom}_{nom}.png'
    img.save(path)
    print(f'{path} saved')
    return path


def get_sheet_list():
    """ Retourne l'ensemble des lignes du GSheets """
    return sh.sheet1.get_all_records()
    

def filtre_name_list():
    """ Retourne la liste des noms / prénoms / email marqués comme à traiter """
    name_list = []
    for elem in sheet_list:
        if elem[c_check] == checkmark:
            current = [elem[c_nom], elem[c_prenom], elem[c_email], sheet_list.index(elem)+2]
            name_list.append(current)

    return name_list


def mark_as_done(row):
    """ Marque la ligne donnée comme traitée """
    sh.sheet1.update(f'G{row[3]}', "sent")
    # for e in sheet_list:
    #     if e[c_nom] == row[0] and e[c_prenom] == row[1]:
    #         e[c_check] = 'Mail envoyé'
    #         break


# Début du programme
config = configparser.ConfigParser()
config.read('config.ini')

diploma_template = config['Mail']['diploma_template_path']

c_nom = config['SheetsAPI']['colonne_nom']
c_prenom = config['SheetsAPI']['colonne_prenom']
c_email = config['SheetsAPI']['colonne_email']
c_check = config['SheetsAPI']['colonne_check']
checkmark = config['SheetsAPI']['checkmark']

sheets_api_credential_json_path = config['SheetsAPI']['sheets_api_credential_json_path']
sheets_file_name = config['SheetsAPI']['sheets_file_name']


# Connect to Sheets API & open file 
gc = gspread.service_account(filename=sheets_api_credential_json_path)
sh = gc.open(sheets_file_name)

sheet_list = get_sheet_list()
name_list = filtre_name_list()
d = str(datetime.datetime.now().replace(microsecond=0))
print(f"Exécution du programme - {d}\n{len(sheet_list)} lignes au total sur le GSheets\n{len(name_list)} lignes marquées comme à traiter\n")

if not os.path.exists('diplomas'):
    os.mkdir('diplomas')
if not os.path.exists('logs'):
    os.mkdir('logs')

success = []
fails = []
for name in name_list:
    try:
        name.append(generate_image(name[0], name[1]))
        send_email(name)
        
        mark_as_done(name)
        success.append(name)
    except Exception as e:
        print(e)
        fails.append(name)

generate_report(success, fails)
