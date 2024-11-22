import cv2
from pyzbar.pyzbar import decode
import numpy as np
import pandas as pd
import time  # Pour gérer les temporisateurs

# Ouvre la webcam (indice 0 pour la caméra par défaut)
cap = cv2.VideoCapture(0)

# Vérifie si la caméra est ouverte
if not cap.isOpened():
    print("Erreur : impossible d'accéder à la caméra")
    exit()

# Chargement des données du fichier Excel (s'il existe) pour garder trace des scans
file_path = 'C:\\Users\\adrie\\Documents\\2MA 175 POL\\Sinterklaas\\qr_code\\Registration.xlsx'

try:
    data = pd.read_excel(file_path)
except FileNotFoundError:
    # Si le fichier n'existe pas, on crée un DataFrame vide avec les colonnes nécessaires
    data = pd.DataFrame(columns=['First name', 'Name', 'Count'])

# Variable pour suivre si un QR code a été détecté
qr_detected = False
last_scan_time = time.time()  # Temps du dernier scan
last_qr_data = ""  # Dernier QR code scanné

# Durée d'attente avant de permettre un nouveau scan (en secondes)
scan_interval = 5  # secondes

while True:
    # Capture une image depuis la caméra
    ret, frame = cap.read()

    # Si l'image est correctement capturée
    if not ret:
        print("Erreur de capture de l'image")
        break
    
    # Décode tous les QR codes dans l'image
    decoded_objects = decode(frame)
    
    # Vérifie si un QR code a été détecté
    if not qr_detected:
        for obj in decoded_objects:
            # Récupère les données du QR code
            qr_data = obj.data.decode('utf-8')
            
            # Ignore les QR codes déjà scannés dans le dernier intervalle
            if qr_data == last_qr_data and time.time() - last_scan_time < scan_interval:
                continue  # Si c'est le même QR code, on passe au suivant

            print(f"QR code détecté : {qr_data}")

            # Supposons que le QR code contient les informations de type "Nom : [nom], Prénom : [prenom]"
            qr_content = qr_data.split(",")
            nom = qr_content[0].split(":")[1].strip()
            prenom = qr_content[1].split(":")[1].strip()

            # Recherche si ce nom et prénom existent déjà dans le DataFrame
            if ((data['First name'] == prenom) & (data['Name'] == nom)).any():
                # Si trouvé, on vérifie la valeur de "Count" et on la modifie si nécessaire
                if data.loc[(data['First name'] == prenom) & (data['Name'] == nom), 'Count'].values[0] == 1:
                    person_info = data.loc[(data['First name'] == prenom) & (data['Name'] == nom)]
                    print(f"Informations de {prenom} {nom}:")
                    print(person_info)
                    # Si "Count" est True (1), on le change en False (0)
                    data.loc[(data['First name'] == prenom) & (data['Name'] == nom), 'Count'] = 0
                    print(f"Premiere fois que {prenom} {nom} est scanné.")
                else:
                    print(f"{prenom} {nom} a deja ete scanne.")
            else:
                # Sinon, on affiche un message pour indiquer que le nom ne correspond pas
                print(f"Aucun enregistrement trouvé pour {prenom} {nom}.")

            # Sauvegarde les données mises à jour dans le fichier Excel immédiatement après chaque modification
            data.to_excel(file_path, index=False)

            # Marquer que le QR code a été détecté
            qr_detected = True
            last_qr_data = qr_data  # Met à jour les données du dernier QR code scanné
            last_scan_time = time.time()  # Met à jour le temps du dernier scan

            # Dessine un rectangle autour du QR code
            rect_points = obj.polygon
            if len(rect_points) == 4:
                pts = [tuple(point) for point in rect_points]
                cv2.polylines(frame, [np.array(pts, dtype=np.int32)], isClosed=True, color=(0, 255, 0), thickness=2)
            else:
                cv2.rectangle(frame, (obj.rect.left, obj.rect.top), (obj.rect.left + obj.rect.width, obj.rect.top + obj.rect.height), (0, 255, 0), 2)

    # Affiche l'image capturée avec les QR codes détectés
    cv2.imshow("QR Code Scanner", frame)

    # Vérifie si le délai de réinitialisation est écoulé pour permettre un nouveau scan
    if qr_detected and time.time() - last_scan_time >= scan_interval:
        qr_detected = False  # Réinitialise la détection après le délai

    # Appuyez sur 'q' pour quitter la boucle
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Libère la caméra et ferme les fenêtres OpenCV
cap.release()
cv2.destroyAllWindows()
