import pandas as pd
import qrcode
import os

# Fonction principale
def generate_qr_codes_from_excel(file_path, output_dir="C:\\Users\\adrie\\Documents\\2MA 175 POL\\Sinterklaas\\qr_code"):
    # Vérifie si le fichier existe
    if not os.path.exists(file_path):
        print(f"Fichier {file_path} introuvable.")
        return
    
    # Crée le dossier de sortie s'il n'existe pas
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Charge les données Excel
    try:
        data = pd.read_excel(file_path)
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier Excel : {e}")
        return
    
    # Vérifie si les colonnes "Nom", "Prénom" et "Prom" existent
    if 'First name' not in data.columns or 'Name' not in data.columns or 'Prom' not in data.columns:
        print("Le fichier Excel doit contenir les colonnes 'Nom', 'Prénom' et 'Prom'.")
        return
    
    # Ajoute une colonne "Count" et remplit chaque élément avec 1
    data['Count'] = 1
    
    # Enregistre les modifications dans le fichier Excel
    data.to_excel(file_path, index=False)
    print("Colonne 'Count' ajoutée et remplie avec 1.")
    
    # Récupère les valeurs uniques de la colonne 'Prom' pour créer des dossiers
    prom_categories = data['Prom'].unique()
    
    # Crée un dossier pour chaque catégorie 'Prom'
    for category in prom_categories:
        category_folder = os.path.join(output_dir, str(category))
        
        # Crée le dossier pour la catégorie, si ce n'est pas déjà fait
        if not os.path.exists(category_folder):
            os.makedirs(category_folder)
            print(f"Dossier créé pour la catégorie 'Prom' : {category}")
        
        # Génère un QR code pour chaque ligne de cette catégorie
        for index, row in data[data['Prom'] == category].iterrows():
            # Récupère le nom, prénom et catégorie
            nom = row['Name']
            prenom = row['First name']
            
            # Texte du QR code : "Nom : [nom], Prénom : [prenom]"
            qr_content = f"Nom : {nom}, Prenom : {prenom}"
            
            # Crée un QR code
            qr = qrcode.QRCode(
                version=1, 
                error_correction=qrcode.constants.ERROR_CORRECT_L, 
                box_size=30, 
                border=4
            )
            qr.add_data(qr_content)
            qr.make(fit=True)
            
            # Crée une image QR code
            img = qr.make_image(fill_color="black", back_color="white")
            
            # Définit le nom du fichier de sortie dans le dossier de la catégorie
            output_file = os.path.join(category_folder, f"QR_{index + 1}_{nom}_{prenom}.png")
            img.save(output_file)
            print(f"QR code généré pour {nom} {prenom} dans la catégorie {category} : {output_file}")
    
    print("Génération des QR codes terminée.")

# Exemple d'utilisation
excel_file = "C:\\Users\\adrie\\Documents\\2MA 175 POL\\Sinterklaas\\qr_code\\Registration.xlsx"  # Chemin vers votre fichier Excel
generate_qr_codes_from_excel(excel_file)
