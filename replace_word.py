import os
from docx import Document

def rechercher_et_remplacer(doc, expressions, remplacements):
    """
    Rechercher et remplacer des expressions dans un document Word.
    """
    modifie = False
    for paragraph in doc.paragraphs:
        for i, expr in enumerate(expressions):
            if expr in paragraph.text:
                paragraph.text = paragraph.text.replace(expr, remplacements[i])
                modifie = True
    return modifie

def modifier_titre(doc, expr_titre, remplacement_titre):
    """
    Modifier le titre d'un document Word.
    """
    modifie = False
    if doc.core_properties.title and expr_titre in doc.core_properties.title:
        doc.core_properties.title = doc.core_properties.title.replace(expr_titre, remplacement_titre)
        modifie = True
    return modifie

def modifier_meta(doc, meta_auteur, meta_modifie_par):
    """
    Modifier les métadonnées "created by" et "modified by" d'un document Word.
    """
    doc.core_properties.author = meta_auteur
    doc.core_properties.last_modified_by = meta_modifie_par

def traiter_fichier(fichier_path, expressions, remplacements, changer_titre, expr_titre, remplacement_titre, changer_meta, meta_auteur, meta_modifie_par):
    """
    Traiter un fichier Word : remplacer les expressions, modifier le titre et les métadonnées.
    """
    try:
        doc = Document(fichier_path)
    except Exception as e:
        # Retourner une erreur si le fichier ne peut pas être ouvert
        return False, fichier_path, str(e)

    modifie = False

    # Remplacer les expressions partout dans le document
    if rechercher_et_remplacer(doc, expressions, remplacements):
        modifie = True

    # Modifier le titre si demandé
    if changer_titre and expr_titre:
        if modifier_titre(doc, expr_titre, remplacement_titre):
            modifie = True

    # Modifier les métadonnées si demandé
    if changer_meta:
        modifier_meta(doc, meta_auteur, meta_modifie_par)

    # Sauvegarder directement dans le fichier original
    if modifie or changer_meta or changer_titre:
        try:
            doc.save(fichier_path)
            print(f"Modifications enregistrées pour : {fichier_path}")
        except Exception as e:
            return False, fichier_path, str(e)

    return modifie, fichier_path, None

def parcourir_repertoire(input_dir, expressions, remplacements, changer_titre, expr_titre, remplacement_titre, changer_meta, meta_auteur, meta_modifie_par, log_path, error_log_path):
    """
    Parcourir les fichiers Word dans un répertoire (et sous-répertoires), appliquer les modifications et écraser les fichiers originaux.
    """
    fichiers_modifies = []
    with open(log_path, 'w') as log_file, open(error_log_path, 'w') as error_log_file:
        for root, _, files in os.walk(input_dir):
            for fichier in files:
                if fichier.endswith('.docx') and not fichier.startswith('~'):  # Ignorer les fichiers temporaires
                    fichier_path = os.path.join(root, fichier)
                    modifie, _, erreur = traiter_fichier(
                        fichier_path, expressions, remplacements, changer_titre, expr_titre, remplacement_titre,
                        changer_meta, meta_auteur, meta_modifie_par
                    )
                    if erreur:
                        error_log_file.write(f"Erreur avec le fichier : {fichier_path}\nErreur : {erreur}\n")
                    elif modifie:
                        fichiers_modifies.append(fichier_path)
                        log_file.write(f"Modifié : {fichier_path}\n")
                    else:
                        log_file.write(f"Non modifié : {fichier_path}\n")
    return fichiers_modifies

def main():
    # Demander le répertoire d'entrée
    print("Entrez le chemin du répertoire contenant les fichiers Word :")
    input_dir = input().strip()

    if not os.path.isdir(input_dir):
        print("Le chemin spécifié n'est pas un répertoire valide.")
        return

    # Demander le nombre d'expressions à changer
    print("Combien d'expressions souhaitez-vous changer ?")
    nb_expressions = int(input().strip())

    # Demander les expressions et leurs remplacements
    expressions = []
    remplacements = []
    for i in range(nb_expressions):
        print(f"Entrez l'expression à remplacer #{i + 1} :")
        expressions.append(input().strip())
        print(f"Entrez le texte de remplacement pour cette expression :")
        remplacements.append(input().strip())

    # Demander si le titre doit être modifié
    print("Souhaitez-vous modifier une expression dans le titre ? (O/N)")
    changer_titre = input().strip().upper() == 'O'
    expr_titre = ""
    remplacement_titre = ""
    if changer_titre:
        print("Entrez l'expression à remplacer dans le titre :")
        expr_titre = input().strip()
        print("Entrez le texte de remplacement pour cette expression dans le titre :")
        remplacement_titre = input().strip()

    # Demander si les métadonnées doivent être modifiées
    print("Souhaitez-vous modifier les métadonnées 'created by' et 'modified by' ? (O/N)")
    changer_meta = input().strip().upper() == 'O'
    meta_auteur = ""
    meta_modifie_par = ""
    if changer_meta:
        print("Entrez le nouveau 'created by' :")
        meta_auteur = input().strip()
        print("Entrez le nouveau 'modified by' :")
        meta_modifie_par = input().strip()

    # Parcourir le répertoire et traiter les fichiers
    log_path = os.path.join(os.getcwd(), "fichiers_modifies_log.txt")
    error_log_path = os.path.join(os.getcwd(), "fichiers_erreurs_log.txt")
    fichiers_modifies = parcourir_repertoire(
        input_dir, expressions, remplacements, changer_titre, expr_titre, remplacement_titre,
        changer_meta, meta_auteur, meta_modifie_par, log_path, error_log_path
    )

    print(f"\nTous les fichiers modifiés ont été écrasés à leur emplacement d'origine.")
    print(f"Un journal des fichiers modifiés est disponible ici : {log_path}")
    print(f"Un journal des erreurs est disponible ici : {error_log_path}")

if __name__ == "__main__":
    main()
