import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import unicodedata
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Image as RLImage
from PIL import Image, ImageTk 
import sys
from datetime import datetime
import locale


# Variables globales
student_file_path = None
admitted_file_path = None

def resource_path(relative_path):
    """ Récupère le chemin absolu du fichier de ressource (image, etc.) """
    try:
        # Si l'application est dans l'exécutable, PyInstaller crée un dossier temporaire
        # où il extrait les fichiers.
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    except Exception as e:
        print(f"Erreur lors de la récupération du fichier : {e}")
        return relative_path

def normalize_name(name):
     # Supprimer les accents et mettre en majuscule
    name = unicodedata.normalize('NFD', name)
    name = ''.join(c for c in name if unicodedata.category(c) != 'Mn')
    return ' '.join(name.upper().split())

# Sélection BD
def select_student_file():
    global student_file_path
    student_file_path = filedialog.askopenfilename(
        title="Sélectionnez le fichier Excel (.xlsx)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if student_file_path:
        filename = os.path.basename(student_file_path)
        student_file_label.config(text=f"Fichier Excel : {filename}")
        student_file_label.tooltip_text = student_file_path
        messagebox.showinfo("Fichier chargé", "Fichier Excel chargé avec succès.")

# Sélection admis
def select_admitted_file():
    global admitted_file_path
    admitted_file_path = filedialog.askopenfilename(
        title="Sélectionnez le fichier des admis (.xlsx)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if admitted_file_path:
        filename = os.path.basename(admitted_file_path)
        admitted_file_label.config(text=f"Fichier Excel : {filename}")
        admitted_file_label.tooltip_text = admitted_file_path
        messagebox.showinfo("Fichier chargé", "Fichier Excel des admis chargé avec succès.")

# Charger les étudiants
def load_student_data(file_path):
    df = pd.read_excel(file_path)
    filename = os.path.basename(file_path)

    num_col = None
    for col in df.columns:
        if str(col).strip().lower() in ["n°", "num", "numéro d'inscription", "numéro", "no", "numero", "n", "numero d'inscription", "n° d'inscription", "n° inscription"]:
            num_col = col
            break
    if not num_col:
        num_col = "Inconnu"  # Valeur par défaut si aucune colonne n'est trouvée
        raise ValueError(f"Assurrez-vous que {filename} contient une colonne pour les numéros d'inscriptions")

     # Nettoyer et normaliser les noms de colonnes
    normalized_cols = [col.strip().lower() for col in df.columns]

    # Ajouter colonne NumInscription AVANT de restreindre les colonnes
    df['NumInscription'] = df[num_col] if num_col != "Inconnu" and num_col in df.columns else "NC"

    # Cas 1 : Vérifier la présence de la colonne "Noms et Prénoms"
    full_name_col = None
    for col in normalized_cols:
        if col in ["noms et prénoms", "nom et prénom", "nom et prenom", "nom et prénom(s)", "nom et prenom(s)", "noms et prénom(s)", "nom et prénoms", "nom et prenoms"]:
            full_name_col = col
            break

    if full_name_col:
        # Si la colonne "Noms et Prénoms" existe, l'utiliser
        df['FullName'] = df[df.columns[normalized_cols.index(full_name_col)]].fillna('').astype(str).str.upper()
    elif 'nom' in normalized_cols and 'prénoms' in normalized_cols:
        noms_col = df.columns[normalized_cols.index('nom')]
        prenoms_col = df.columns[normalized_cols.index('prénoms')]
        
        df[noms_col] = df[noms_col].fillna('').astype(str).str.upper()
        df[prenoms_col] = df[prenoms_col].fillna('').astype(str).str.upper()
        
        df['FullName'] = df.apply(
            lambda row: f"{row[noms_col]} {row[prenoms_col]}".strip(), axis=1
        )

    elif 'noms' in normalized_cols and 'prénoms' in normalized_cols:
        noms_col = df.columns[normalized_cols.index('noms')]
        prenoms_col = df.columns[normalized_cols.index('prénoms')]
        
        df[noms_col] = df[noms_col].fillna('').astype(str).str.upper()
        df[prenoms_col] = df[prenoms_col].fillna('').astype(str).str.upper()
        
        df['FullName'] = df.apply(
            lambda row: f"{row[noms_col]} {row[prenoms_col]}".strip(), axis=1
        )
    else:
        raise ValueError(f"Assurrez-vous que {filename} contient une ou deux colonne(s) pour les noms et prénoms")

    student_set = set(df[['FullName', 'NumInscription']].itertuples(index=False, name=None))

    return student_set

# Charger les admis avec ajout du numéro d'inscription
def load_admitted_students(file_path):
    df = pd.read_excel(file_path)
    filename = os.path.basename(file_path)

    # Identifier automatiquement la colonne du numéro d'inscription
    num_col = None
    for col in df.columns:
        if str(col).strip().lower() in ["n°", "num", "numéro d'inscription", "numéro", "no", "numero", "n", "num", "numero d'inscription", "n° d'inscription", "n° inscription"]:
            num_col = col
            break
    if not num_col:
        num_col = "Inconnu"  # Valeur par défaut si aucune colonne n'est trouvée
        raise ValueError(f"Assurrez-vous que {filename} contient une colonne pour les numéros d'inscriptions")

    # Nettoyer et normaliser les noms de colonnes
    normalized_cols = [col.strip().lower() for col in df.columns]

    # Cas 1 : Vérifier la présence de la colonne "Noms et Prénoms"
    full_name_col = None
    for col in normalized_cols:
        if col in ["noms et prénoms", "nom et prénom", "nom et prenom", "nom et prénom(s)", "nom et prenom(s)", "noms et prénom(s)", "nom et prénoms", "nom et prenoms"]:
            full_name_col = col
            break

    if full_name_col:
        # Si la colonne "Noms et Prénoms" existe, l'utiliser
        df['FullName'] = df[df.columns[normalized_cols.index(full_name_col)]].fillna('').astype(str).str.upper()
    elif 'nom' in normalized_cols and 'prénoms' in normalized_cols:
        noms_col = df.columns[normalized_cols.index('nom')]
        prenoms_col = df.columns[normalized_cols.index('prénoms')]
        
        df[noms_col] = df[noms_col].fillna('').astype(str).str.upper()
        df[prenoms_col] = df[prenoms_col].fillna('').astype(str).str.upper()
        
        df['FullName'] = df.apply(
            lambda row: f"{row[noms_col]} {row[prenoms_col]}".strip(), axis=1
        )

    elif 'noms' in normalized_cols and 'prénoms' in normalized_cols:
        noms_col = df.columns[normalized_cols.index('noms')]
        prenoms_col = df.columns[normalized_cols.index('prénoms')]
        
        df[noms_col] = df[noms_col].fillna('').astype(str).str.upper()
        df[prenoms_col] = df[prenoms_col].fillna('').astype(str).str.upper()
        
        df['FullName'] = df.apply(
            lambda row: f"{row[noms_col]} {row[prenoms_col]}".strip(), axis=1
        )
    elif 'noms' in normalized_cols and 'prénom' in normalized_cols:
        noms_col = df.columns[normalized_cols.index('noms')]
        prenoms_col = df.columns[normalized_cols.index('prénom')]
        
        df[noms_col] = df[noms_col].fillna('').astype(str).str.upper()
        df[prenoms_col] = df[prenoms_col].fillna('').astype(str).str.upper()
        
        df['FullName'] = df.apply(
            lambda row: f"{row[noms_col]} {row[prenoms_col]}".strip(), axis=1
        )
    elif 'nom' in normalized_cols and 'prénom' in normalized_cols:
        noms_col = df.columns[normalized_cols.index('nom')]
        prenoms_col = df.columns[normalized_cols.index('prénom')]
        
        df[noms_col] = df[noms_col].fillna('').astype(str).str.upper()
        df[prenoms_col] = df[prenoms_col].fillna('').astype(str).str.upper()
        
        df['FullName'] = df.apply(
            lambda row: f"{row[noms_col]} {row[prenoms_col]}".strip(), axis=1
        )
    else:
        raise ValueError(f"Assurrez-vous que {filename} contient une ou deux colonnes pour les noms et prénoms")

    # Appliquer la normalisation des noms
    df['FullName'] = df['FullName'].apply(normalize_name)

    # Ajouter la colonne numéro d'inscription
    df['NumInscription'] = df[num_col] if num_col != "Inconnu" else "NC"

    # Retourner les données sous forme de liste ordonnée de tuples
    admitted_list = [(row['FullName'], row['NumInscription']) for _, row in df.iterrows()]
    return admitted_list

def compare_students(mes_etudiants, admitted_students):
    """
    Compare tous les étudiants avec la liste officielle des admis.
    Utilise NumInscription comme clé unique.
    Retourne un DataFrame avec tous les étudiants (admis et non admis),
    et trie les admis selon l’ordre d’apparition dans la liste officielle.
    """
    # Dictionnaire rapide : NumInscription => Nom
    etudiants_dict = {insc: nom for nom, insc in mes_etudiants}

    # Ensemble des NumInscription admis
    admitted_ids = {num_insc for _, num_insc in admitted_students}

    # Création d'une liste ordonnée pour extraire l'ordre des admis
    ordre_admis = [num_insc for _, num_insc in admitted_students]

    # Construire la liste complète avec True/False
    resultat = []
    for num_insc, nom in etudiants_dict.items():
        admis = num_insc in admitted_ids
        resultat.append((nom, num_insc, admis))

    # Création du DataFrame final
    df_resultat = pd.DataFrame(resultat, columns=["FullName", "NumInscription", "Admis"])

    # Pour les admis, ajouter une colonne Order selon leur position dans la liste officielle
    df_resultat["Order"] = df_resultat.apply(
        lambda row: ordre_admis.index(row["NumInscription"]) if row["Admis"] and row["NumInscription"] in ordre_admis else -1,
        axis=1
    )

    return df_resultat

# Calcul pourcentage
def calculate_success_percentage(df):
    total = len(df)
    admitted = df['Admis'].sum()
    return (admitted / total) * 100, admitted, total



# Traitement principal
def process_files():
    try:
        if not student_file_path or not admitted_file_path:
            messagebox.showwarning("Attention", "Veuillez sélectionner les deux fichiers avant de continuer.")
            return

        # Charger les données des étudiants et des admis
        students_df = load_student_data(student_file_path)
        admitted_data = load_admitted_students(admitted_file_path)

        # Comparer les étudiants et déterminer ceux qui sont admis
        students_df = compare_students(students_df, admitted_data)

        # Calculer le pourcentage de réussite et les statistiques
        success_percentage, admitted_count, total_count = calculate_success_percentage(students_df)

        # Filtrer les étudiants admis et trier par ordre de mérite (colonne "Order")
        admitted_only_df = students_df[students_df['Admis'] == True].copy()
        admitted_only_df = admitted_only_df.sort_values("Order")  # ✅ TRI ICI

        # Nettoyage du tableau dans l'interface
        for item in admitted_table.get_children():
            admitted_table.delete(item)

        # Affichage des étudiants admis dans le tableau avec zebra (alternance des lignes)
        for display_idx, (_, row) in enumerate(admitted_only_df.iterrows()):
            full_name = row['FullName']
            num_inscription = row['NumInscription']
            tag = 'evenrow' if display_idx % 2 == 0 else 'oddrow'
            admitted_table.insert('', tk.END, values=(num_inscription, full_name), tags=(tag,))

        # Affichage des résultats dans la section de résultats visuels
        for widget in result_frame.winfo_children():
            widget.destroy()

        # Masquer le slogan et le copyright
        splash_frame.pack_forget()


        if admitted_count > 0:
            if success_percentage > 80:
                color = "#20BF6B"  # vert
            elif success_percentage >= 40:
                color = "#FFA500"  # orange
            else:
                color = "#FF4C4C"  # rouge

            tk.Label(
                result_frame,
                text=f"🎉 Félicitations aux {admitted_count}/{total_count} candidats admis, inscrits au Cours Prépa T&R!",
                font=("Arial", 12)
            ).pack(pady=5)

            tk.Label(
                result_frame,
                text=f"Taux de réussite : {success_percentage:.2f}%",
                font=("Arial", 14, "bold"),
                fg=color
            ).pack(pady=5)

            save_button.pack(side=tk.LEFT, padx=10)
            show_non_admitted_button.pack(side=tk.LEFT, padx=10)
        else:
            tk.Label(
                result_frame,
                text="😔 Aucun étudiant admis cette fois.\nNe baissez pas les bras, continuez vos efforts ! 💪",
                font=("Arial", 14),
                fg="#FF4C4C"
            ).pack(pady=5)
            save_button.pack_forget()
            show_non_admitted_button.pack_forget()

    except Exception as e:
        messagebox.showerror("Erreur", str(e))
# Traitement principal
def process_files():
    try:
        if not student_file_path or not admitted_file_path:
            messagebox.showwarning("Attention", "Veuillez sélectionner les deux fichiers avant de continuer.")
            return

        # Charger les données des étudiants et des admis
        students_df = load_student_data(student_file_path)
        admitted_data = load_admitted_students(admitted_file_path)

        # Comparer les étudiants et déterminer ceux qui sont admis
        students_df = compare_students(students_df, admitted_data)

        # Calculer le pourcentage de réussite et les statistiques
        success_percentage, admitted_count, total_count = calculate_success_percentage(students_df)

        # Filtrer les étudiants admis et trier par ordre de mérite (colonne "Order")
        admitted_only_df = students_df[students_df['Admis'] == True].copy()
        admitted_only_df = admitted_only_df.sort_values("Order")  # ✅ TRI ICI

        # Nettoyage du tableau dans l'interface
        for item in admitted_table.get_children():
            admitted_table.delete(item)

        # Affichage des étudiants admis dans le tableau avec zebra (alternance des lignes)
        for display_idx, (_, row) in enumerate(admitted_only_df.iterrows()):
            full_name = row['FullName']
            num_inscription = row['NumInscription']
            tag = 'evenrow' if display_idx % 2 == 0 else 'oddrow'
            admitted_table.insert('', tk.END, values=(num_inscription, full_name), tags=(tag,))

        # Affichage des résultats dans la section de résultats visuels
        for widget in result_frame.winfo_children():
            widget.destroy()

        # Masquer le slogan et le copyright
        splash_frame.pack_forget()


        if admitted_count > 0:
            if success_percentage > 80:
                color = "#20BF6B"  # vert
            elif success_percentage >= 40:
                color = "#FFA500"  # orange
            else:
                color = "#FF4C4C"  # rouge

            tk.Label(
                result_frame,
                text=f"🎉 Félicitations aux {admitted_count}/{total_count} candidats admis, inscrits au Cours Prépa T&R!",
                font=("Arial", 12)
            ).pack(pady=5)

            tk.Label(
                result_frame,
                text=f"Taux de réussite : {success_percentage:.2f}%",
                font=("Arial", 14, "bold"),
                fg=color
            ).pack(pady=5)

            save_button.pack(side=tk.LEFT, padx=10)
            show_non_admitted_button.pack(side=tk.LEFT, padx=10)
        else:
            tk.Label(
                result_frame,
                text="😔 Aucun étudiant admis cette fois.\nNe baissez pas les bras, continuez vos efforts ! 💪",
                font=("Arial", 14),
                fg="#FF4C4C"
            ).pack(pady=5)
            save_button.pack_forget()
            show_non_admitted_button.pack_forget()

    except Exception as e:
        messagebox.showerror("Erreur", str(e))



# Afficher les étudiants non admis
def display_non_admitted():
    try:
        if not student_file_path or not admitted_file_path:
            messagebox.showwarning("Attention", "Veuillez sélectionner les deux fichiers avant de continuer.")
            return

        students_df = load_student_data(student_file_path)
        admitted_students = load_admitted_students(admitted_file_path)

        # Comparer les étudiants et déterminer leur statut d'admission
        students_df = compare_students(students_df, admitted_students)

        # Filtrer les étudiants non admis
        non_admitted_df = students_df[students_df['Admis'] == False]

        # Nettoyer le tableau dans l'interface
        for item in admitted_table.get_children():
            admitted_table.delete(item)

        # Affichage des étudiants admis dans le tableau avec zebra (alternance des lignes)
        for display_idx, (_, row) in enumerate(non_admitted_df.iterrows()):
            full_name = row['FullName']
            num_inscription = row['NumInscription']
            tag = 'evenrow' if display_idx % 2 == 0 else 'oddrow'
            admitted_table.insert('', tk.END, values=(num_inscription, full_name), tags=(tag,))

        # Changer le texte du bouton pour afficher la liste des admis
        show_non_admitted_button.config(text="ADMIS", command=display_admitted)


    except Exception as e:
        messagebox.showerror("Erreur", str(e))


# Afficher les étudiants admis
def display_admitted():
    try:
        if not student_file_path or not admitted_file_path:
            messagebox.showwarning("Attention", "Veuillez sélectionner les deux fichiers avant de continuer.")
            return

        # Charger les données des étudiants et des admis
        students_df = load_student_data(student_file_path)
        admitted_data = load_admitted_students(admitted_file_path)

        # Comparer les étudiants et déterminer ceux qui sont admis
        students_df = compare_students(students_df, admitted_data)

        # Trier les étudiants admis par ordre d'apparition dans la liste des admis
        admitted_only_df = students_df[students_df['Admis'] == True].copy()
        # Ajouter une colonne "Order" pour spécifier l'ordre de mérite selon la liste des admis
        admitted_only_df = admitted_only_df.sort_values("Order")


        # Nettoyage du tableau dans l'interface
        for item in admitted_table.get_children():
            admitted_table.delete(item)

        # Affichage des étudiants admis dans le tableau avec zebra (alternance des lignes)
        for display_idx, (_, row) in enumerate(admitted_only_df.iterrows()):
            full_name = row['FullName']
            num_inscription = row['NumInscription']
            tag = 'evenrow' if display_idx % 2 == 0 else 'oddrow'
            admitted_table.insert('', tk.END, values=(num_inscription, full_name), tags=(tag,))

        # Changer le texte du bouton pour afficher la liste des non admis
        show_non_admitted_button.config(text="NON ADMIS", command=display_non_admitted)

    except Exception as e:
        messagebox.showerror("Erreur", str(e))



# Export pdf

def save_admitted_to_pdf():
    try:
        if not student_file_path or not admitted_file_path:
            messagebox.showwarning("Attention", "Veuillez d'abord traiter les fichiers.")
            return

        # Charger les données
        students_df = load_student_data(student_file_path)
        admitted_students = load_admitted_students(admitted_file_path)

        # Comparer
        students_df = compare_students(students_df, admitted_students)
        admitted_df = students_df[students_df['Admis'] == True].copy()

        # Trier les étudiants par ordre de mérite (ou d'apparition dans la liste des admis)
        if admitted_df.empty:
            messagebox.showinfo("Information", "Aucun étudiant admis à sauvegarder.")
            return

        admitted_df = admitted_df.sort_values("Order")

        # Chemin de sauvegarde
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Fichiers PDF", "*.pdf")],
            title="Enregistrer sous..."
        )

        if save_path:
            # Création du PDF
            doc = SimpleDocTemplate(
                save_path,
                pagesize=A4,
                topMargin=1*cm,
                leftMargin=2*cm,
                rightMargin=2*cm,
                bottomMargin=2*cm
            )

            elements = []
            styles = getSampleStyleSheet()

            # Ajouter le logo
            logo_path = resource_path("logo/logo_cours.jpg")
            if os.path.exists(logo_path):
                logo = RLImage(logo_path, width=3*cm, height=3*cm)
                logo.hAlign = "CENTER"
                elements.append(logo)
            
            # Définir la locale en français pour afficher le mois en français
            locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

            # Ajouter la date après le logo
            date_str = "Antananarivo, le " + datetime.now().strftime("%d %B %Y")

            # Style de la date
            date_style = styles["Normal"]
            date_style.fontName = "Helvetica"
            date_style.fontSize = 10
            date_style.alignment = 1  # Centré

            # Créer le paragraphe avec la date
            date = Paragraph(date_str, date_style)
            elements.append(date)

            # Ajouter un petit espacement après la date
            elements.append(Spacer(1, 12))
            # Ajouter le titre
            title_style = styles["Normal"]
            title_style.fontName = "Helvetica-Bold"
            title = Paragraph("LISTE DE NOS CANDIDATS ADMIS AUX CONCOURS", title_style)
            elements.append(title)
            elements.append(Spacer(1, 12))

            # Préparer les données du tableau
            admitted_df = admitted_df[['NumInscription', 'FullName']].copy()
            admitted_df.columns = ['NUMÉRO D\'INSCRIPTION', 'NOM ET PRÉNOMS']
            data = [admitted_df.columns.tolist()] + admitted_df.values.tolist()

            col_width = [6 * cm, 10 * cm]
            table = Table(data, colWidths=col_width, repeatRows=1)
            table.setStyle(TableStyle([
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ]))

            elements.append(table)
            doc.build(elements)

            messagebox.showinfo("Succès", f"PDF enregistré avec succès à :\n{save_path}")

    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde PDF : {str(e)}")
def save_admitted_to_pdf():
    try:
        if not student_file_path or not admitted_file_path:
            messagebox.showwarning("Attention", "Veuillez d'abord traiter les fichiers.")
            return

        # Charger les données
        students_df = load_student_data(student_file_path)
        admitted_students = load_admitted_students(admitted_file_path)

        # Comparer
        students_df = compare_students(students_df, admitted_students)
        admitted_df = students_df[students_df['Admis'] == True].copy()

        # Trier les étudiants par ordre de mérite (ou d'apparition dans la liste des admis)
        if admitted_df.empty:
            messagebox.showinfo("Information", "Aucun étudiant admis à sauvegarder.")
            return

        admitted_df = admitted_df.sort_values("Order")

        # Chemin de sauvegarde
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Fichiers PDF", "*.pdf")],
            title="Enregistrer sous..."
        )

        if save_path:
            # Création du PDF
            doc = SimpleDocTemplate(
                save_path,
                pagesize=A4,
                topMargin=1*cm,
                leftMargin=2*cm,
                rightMargin=2*cm,
                bottomMargin=2*cm
            )

            elements = []
            styles = getSampleStyleSheet()

            # Ajouter le logo
            logo_path = resource_path("logo/logo_cours.jpg")
            if os.path.exists(logo_path):
                logo = RLImage(logo_path, width=3*cm, height=3*cm)
                logo.hAlign = "CENTER"
                elements.append(logo)
            
            # Définir la locale en français pour afficher le mois en français
            locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

            # Ajouter la date après le logo
            date_str = "Antananarivo, le " + datetime.now().strftime("%d %B %Y")

            # Style de la date
            date_style = styles["Normal"]
            date_style.fontName = "Helvetica"
            date_style.fontSize = 10
            date_style.alignment = 1  # Centré

            # Créer le paragraphe avec la date
            date = Paragraph(date_str, date_style)
            elements.append(date)

            # Ajouter un petit espacement après la date
            elements.append(Spacer(1, 12))
            # Ajouter le titre
            title_style = styles["Normal"]
            title_style.fontName = "Helvetica-Bold"
            title = Paragraph("LISTE DE NOS CANDIDATS ADMIS AUX CONCOURS", title_style)
            elements.append(title)
            elements.append(Spacer(1, 12))

            # Préparer les données du tableau
            admitted_df = admitted_df[['NumInscription', 'FullName']].copy()
            admitted_df.columns = ['NUMÉRO D\'INSCRIPTION', 'NOM ET PRÉNOMS']
            data = [admitted_df.columns.tolist()] + admitted_df.values.tolist()

            col_width = [6 * cm, 10 * cm]
            table = Table(data, colWidths=col_width, repeatRows=1)
            table.setStyle(TableStyle([
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ]))

            elements.append(table)
            doc.build(elements)

            messagebox.showinfo("Succès", f"PDF enregistré avec succès à :\n{save_path}")

    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde PDF : {str(e)}")

# Gestion des tooltips (infos-bulles)
def show_tooltip(event):
    widget = event.widget
    tooltip_text = getattr(widget, 'tooltip_text', None)
    if tooltip_text:
        tooltip_label.config(text=tooltip_text)
        tooltip_label.place(x=event.x_root - root.winfo_rootx() + 10, y=event.y_root - root.winfo_rooty() + 10)

def hide_tooltip(event):
    tooltip_label.place_forget()

# Interface utilisateur
root = tk.Tk()
root.title("Cours Prépa")
root.state('zoomed')

screen_height = root.winfo_screenheight()
treeview_height_pixels = int(screen_height * 0.6)

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill='both', expand=True)

button_frame = tk.Frame(frame)
button_frame.pack(pady=(30,5))

load_excel_button = tk.Button(
    button_frame, text="LISTE DE NOS CANDIDATS", command=select_student_file,
    bg="#4B7BEC", fg="white", font=("Arial", 10, "bold"),
    relief="flat", padx=20, pady=10, activebackground="#4B7BEC", cursor="hand2"
)
load_excel_button.pack(side=tk.LEFT, padx=10)

load_txt_button = tk.Button(
    button_frame, text="RESULTAT OFFICIEL", command=select_admitted_file,
    bg="#4B7BEC", fg="white", font=("Arial", 10, "bold"),
    relief="flat", padx=20, pady=10, activebackground="#4B7BEC", cursor="hand2"
)
load_txt_button.pack(side=tk.LEFT, padx=10)

process_button = tk.Button(
    button_frame, text="RESULTAT", command=process_files,
    bg="#20BF6B", fg="white", font=("Arial", 10, "bold"),
    relief="flat", padx=20, pady=10, activebackground="#20BF6B", cursor="hand2"
)
process_button.pack(side=tk.LEFT, padx=10)

save_button = tk.Button(
    button_frame, text="PDF", command=save_admitted_to_pdf,
    bg="#4B7BEC", fg="white", font=("Arial", 10, "bold"),
    relief="flat", padx=20, pady=10, activebackground="#4B7BEC", cursor="hand2"
)
save_button.pack(side=tk.LEFT, padx=10)
save_button.pack_forget()

# Bouton pour afficher les non admis
show_non_admitted_button = tk.Button(
    button_frame, text="NON ADMIS", command=display_non_admitted,
    bg="#FF4C4C", fg="white", font=("Arial", 10, "bold"),
    relief="flat", padx=20, pady=10, activebackground="#FF4C4C", cursor="hand2"
)
show_non_admitted_button.pack(side=tk.LEFT, padx=10)
show_non_admitted_button.pack_forget()  # Cache ce bouton par défaut

# Labels pour afficher les fichiers chargés
student_file_label = tk.Label(frame, text="Aucune liste des candidats chargé.", font=("Courier New", 10))
student_file_label.pack()
student_file_label.bind("<Enter>", show_tooltip)
student_file_label.bind("<Leave>", hide_tooltip)

admitted_file_label = tk.Label(frame, text="Aucune liste des admis chargé.", font=("Courier New", 10))
admitted_file_label.pack()
admitted_file_label.bind("<Enter>", show_tooltip)
admitted_file_label.bind("<Leave>", hide_tooltip)

# Tooltip caché par défaut
tooltip_label = tk.Label(root, text="", bg="yellow", relief="solid", bd=1, font=("Courier New", 9))
tooltip_label.place_forget()

# Tableau des résultats
table_frame = tk.Frame(frame)
table_frame.pack(pady=10, fill='both', expand=True)

style = ttk.Style()
style.configure("Custom.Treeview", font=("Arial", 10), rowheight=30,
                background="#F0F0F0", fieldbackground="#F0F0F0", foreground="#000000")
style.configure("Custom.Treeview.Heading", font=("Arial", 10, "bold"), foreground="#000000")
style.map("Custom.Treeview", background=[('selected', '#20BF6B')], foreground=[('selected', 'white')])

admitted_table = ttk.Treeview(table_frame, columns=("NumeroInscription", "NomPrenom"), show="headings", style="Custom.Treeview")
admitted_table.heading("NumeroInscription", text="NUMERO D'INSCRIPTION")
admitted_table.heading("NomPrenom", text="NOM ET PRENOMS")
admitted_table.column("NumeroInscription", width=80, anchor='center')
admitted_table.column("NomPrenom", width=600, anchor='w')
admitted_table.pack(side=tk.LEFT, fill='both', expand=True)

# Appliquer le style zébré
admitted_table.tag_configure('evenrow', background="#F5F5F5")  # ligne paire
admitted_table.tag_configure('oddrow', background="#FFFFFF")   # ligne impaire

scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=admitted_table.yview)
admitted_table.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill="y")

# Charger le logo
logo_path = resource_path("logo/app.png")  # Assure-toi que ce chemin est correct
if os.path.exists(logo_path):
    logo_img = Image.open(logo_path)
    logo_img = logo_img.resize((100, 100))  # Redimensionne selon ton besoin
    logo_tk = ImageTk.PhotoImage(logo_img)

    # Ajouter le logo dans un label
    logo_label = tk.Label(root, image=logo_tk, borderwidth=0)
    logo_label.image = logo_tk  # Conserver une référence

    logo_label.place(relx=1.0, rely=0.0, anchor="ne", x=-100, y=10)

# Slogan et copyright placés comme un vrai footer centré en bas
splash_frame = tk.Frame(root)
splash_frame.pack(side="bottom",pady=15)

slogan_label = tk.Label(splash_frame, text="Higher and higher!", font=("Arial", 14), fg="#999999")
slogan_label.pack()

copyright_label = tk.Label(
    splash_frame, text="© 2025 Cours Prépa", font=("Arial", 10), fg="#999999"
)
copyright_label.pack()



result_frame = tk.Frame(frame)
result_frame.pack(pady=20)

root.mainloop()