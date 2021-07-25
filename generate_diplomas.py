import gspread, os, datetime, yaml
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


def generate_image(specs, nom, prenom):
    """ Enregistre l'attestation avec le nom et prénom inscrit """

    img = Image.open(specs['diploma_path'])
    draw = ImageDraw.Draw(img)
    nom = nom.upper()
    prenom = prenom.upper()

    txt = f'{prenom} {nom}'
    frame = specs['frame_width']
    fontsize = specs['fontsize']

    # Check if text width exceeds the frame width
    while True:
        font = ImageFont.truetype(specs['font'], fontsize)
        txt_width = get_text_dimensions(txt, font)[0]
        if txt_width < frame:
            break
        fontsize -= 10

    # Calculate text x position and draw text
    w_max = specs['width']
    perfect = (w_max - txt_width) / 2
    draw.text((perfect, specs['text_ypos']), txt, (255, 255, 255), font=font)

    path = f'diplomas/{prenom}_{nom}.png'
    img.save(path)
    print(f'{path} saved')
    return path


def get_sheet_list():
    """ Retourne l'ensemble des lignes du GSheets """
    return sh.sheet1.get_all_records()
    

def filter_name_list(specs):
    """ Retourne la liste des noms / prénoms / email marqués comme à traiter """
    name_list = []
    for elem in sheet_list:
        if elem[specs['_check']] == specs['checkmark']:
            current = [elem[specs['_nom']], elem[specs['_prenom']], elem[specs['_email']], sheet_list.index(elem)+2]
            name_list.append(current)

    return name_list


def mark_as_done(row):
    """ Marque la ligne donnée comme traitée """
    sh.sheet1.update(f'G{row[3]}', "Attestation envoyée !")


# Début du programme
if __name__=='__main__':
    # Read config file
    with open("config.yaml", "r") as rf:
        config = (yaml.load(rf, Loader=yaml.Loader))

    # Connect to Sheets API & open file 
    gc = gspread.service_account(filename=config['Sheets']['credential_path'])
    sh = gc.open(config['Sheets']['file_name'])

    sheet_list = get_sheet_list()
    name_list = filter_name_list(config['Sheet_specs'])

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
            name.append(generate_image(config['Diploma_specs'], name[0], name[1]))
            send_email(name)
            
            mark_as_done(name)
            success.append(name)
        except Exception as e:
            print(e)
            fails.append(name)

    generate_report(success, fails)
    print(f"Fin de l'exécution, le fichier `logs/log_{d}.log` a été généré.")
