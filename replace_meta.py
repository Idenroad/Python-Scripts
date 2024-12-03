import os
import subprocess
from docx import Document
import pikepdf

def convertir_doc_en_docx(fichier_path):
    """
    Convertir un fichier .doc en .docx en utilisant LibreOffice.
    """
    try:
        # Construire le chemin de sortie pour le fichier .docx
        nouveau_chemin = os.path.splitext(fichier_path)[0] + ".docx"

        # Utiliser LibreOffice pour convertir le fichier
        subprocess.run(["libreoffice", "--headless", "--convert-to", "docx", fichier_path, "--outdir", os.path.dirname(fichier_path)], check=True)

        print(f"Converti : {fichier_path} -> {nouveau_chemin}")
        return nouveau_chemin
    except subprocess.CalledProcessError as e:
        print(f"Erreur lors de la conversion de {fichier_path} : {e}")
        return None

def modifier_meta_docx(fichier_path):
    """
    Modifier les métadonnées d'un fichier Word (.docx).
    """
    try:
        doc = Document(fichier_path)
        # Modifier les métadonnées
        doc.core_properties.author = "META_A_METTRE"
        doc.core_properties.last_modified_by = "META_A_METTRE"
        # Sauvegarder les modifications
        doc.save(fichier_path)
        print(f"Métadonnées modifiées pour : {fichier_path}")
    except Exception as e:
        print(f"Erreur lors de la modification des métadonnées pour {fichier_path} : {e}")

def modifier_meta_pdf(fichier_path):
    """
    Modifier les métadonnées d'un fichier PDF.
    """
    try:
        # Ouvrir le fichier PDF avec autorisation d'écraser l'entrée
        with pikepdf.open(fichier_path, allow_overwriting_input=True) as pdf:
            # Modifier les métadonnées DocumentInfo
            pdf.docinfo["/Author"] = "META_A_METTRE"
            pdf.docinfo["/Producer"] = "META_A_METTRE"
            pdf.docinfo["/Creator"] = "META_A_METTRE"

            # Modifier les métadonnées XMP (si disponibles)
            try:
                with pdf.open_metadata(set_pikepdf_as_editor=True) as meta:
                    meta["dc:creator"] = ["META_A_METTRE"]  # Doit être une liste
                    meta["xmp:CreatorTool"] = "META_A_METTRE"
                    meta["pdf:Producer"] = "META_A_METTRE"
            except Exception as e:
                print(f"Impossible de modifier les métadonnées XMP pour {fichier_path} : {e}")

            # Sauvegarder les modifications
            pdf.save(fichier_path)
        print(f"Métadonnées modifiées pour : {fichier_path}")
    except Exception as e:
        print(f"Erreur lors de la modification des métadonnées pour {fichier_path} : {e}")

def parcourir_repertoire_et_modifier_meta(repertoire):
    """
    Parcourir un répertoire et ses sous-répertoires pour modifier les métadonnées
    des fichiers Word (.docx et .doc) et PDF.
    """
    for root, _, files in os.walk(repertoire):
        for fichier in files:
            fichier_path = os.path.join(root, fichier)
            if fichier.endswith(".docx"):
                modifier_meta_docx(fichier_path)
            elif fichier.endswith(".doc"):
                # Convertir en .docx puis modifier les métadonnées
                nouveau_fichier = convertir_doc_en_docx(fichier_path)
                if nouveau_fichier:
                    modifier_meta_docx(nouveau_fichier)
            elif fichier.endswith(".pdf"):
                modifier_meta_pdf(fichier_path)

def main():
    # Demander le répertoire à l'utilisateur
    print("Entrez le chemin du répertoire à traiter :")
    repertoire = input().strip()

    # Vérifier si le répertoire existe
    if not os.path.isdir(repertoire):
        print("Le chemin spécifié n'est pas un répertoire valide.")
        return

    # Parcourir le répertoire et modifier les métadonnées
    parcourir_repertoire_et_modifier_meta(repertoire)
    print("\nTraitement terminé.")

if __name__ == "__main__":
    main()
