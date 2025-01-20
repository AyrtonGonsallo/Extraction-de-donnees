import tkinter as tk
from re import split
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import PyPDF2
import re
import csv

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Application de Traitement de Fichiers")
        self.geometry("700x500")
        self.resultats = []
        # Variables
        self.filepath = ""
        self.totalpages = 0
        self.export_format = tk.StringVar(value="csv")

        # Zone de texte pour afficher les résultats
        self.txt_results = scrolledtext.ScrolledText(self, width=50, height=15, bg='black', fg='white', insertbackground='white')
        self.txt_results.place(relx=0.65, rely=0.15, anchor="n")

        # Cadre pour les boutons
        self.frame_actions = tk.Frame(self)
        self.frame_actions.place(relx=0.2, rely=0.5, anchor="center")

        btn_style = {"width": 20, "bg": "black", "fg": "white"}

        self.btn_import = tk.Button(self.frame_actions, text="Importer un fichier", command=self.import_file, **btn_style)
        self.btn_import.pack(pady=5)

        self.btn_process = tk.Button(self.frame_actions, text="Traiter le fichier", command=self.process_file, state=tk.DISABLED, **btn_style)
        self.btn_process.pack(pady=5)

        self.btn_export = tk.Button(self.frame_actions, text="Exporter le fichier", command=self.export_file, state=tk.DISABLED, **btn_style)
        self.btn_export.pack(pady=5)

        self.btn_clear = tk.Button(self.frame_actions, text="Effacer tout", command=self.clear_all, state=tk.DISABLED, **btn_style)
        self.btn_clear.pack(pady=5)

        # Cadre pour les options d'exportation
        self.frame_format = tk.Frame(self)
        self.frame_format.place(relx=0.2, rely=0.75, anchor="center")

        self.lbl_format = tk.Label(self.frame_format, text="Format d'exportation :")
        self.lbl_format.pack(anchor="w")

        self.radio_csv = tk.Radiobutton(self.frame_format, text="CSV", variable=self.export_format, value="csv")
        self.radio_csv.pack(anchor="w")

        self.radio_xlsx = tk.Radiobutton(self.frame_format, text="XLSX", variable=self.export_format, value="xlsx")
        self.radio_xlsx.pack(anchor="w")

    def import_file(self):
        self.filepath = filedialog.askopenfilename(
            filetypes=[("Fichiers PDF", "*.pdf")]
        )
        if self.filepath:
            try:
                with open(self.filepath, 'rb') as fichier_pdf:
                    lecteur_pdf = PyPDF2.PdfReader(fichier_pdf)
                    nombre_pages = len(lecteur_pdf.pages)
                    self.totalpages=nombre_pages
                    print(f"Le fichier sélectionné contient {nombre_pages} pages.")

                self.txt_results.insert(tk.END, "Le fichier sélectionné contient "+str(nombre_pages)+" pages.\n")
                self.txt_results.insert(tk.END, "fichier " + self.filepath + " lu\n")
                self.btn_process.config(state=tk.NORMAL)
                self.btn_clear.config(state=tk.NORMAL)
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier: {e}")

    def process_file(self):
        for i in range(self.totalpages):
            res1=self.lire_page_pdf(self.filepath,i)
            self.txt_results.insert(tk.END, "Page "+str(i)+"\n")
            res=self.extraire_chaine(res1)
            for re in res:
                self.txt_results.insert(tk.END,re+'\n')
                self.resultats.append(re)
            self.txt_results.insert(tk.END, " "*15)
            self.txt_results.insert(tk.END, "-"*12 )
            self.txt_results.insert(tk.END, '\n')



        self.txt_results.insert(tk.END, "\n")
        self.btn_export.config(state=tk.NORMAL)

    def lire_page_pdf(self,chemin_fichier, numero_page):
        try:
            # Ouvrir le fichier PDF en mode lecture binaire
            with open(chemin_fichier, 'rb') as fichier_pdf:
                lecteur_pdf = PyPDF2.PdfReader(fichier_pdf)

                # Vérifier que le numéro de page est valide
                if numero_page < 0 or numero_page >= len(lecteur_pdf.pages):
                    return f"Numéro de page invalide. Le PDF contient {len(lecteur_pdf.pages)} pages."

                # Obtenir la page spécifiée
                page = lecteur_pdf.pages[numero_page]
                texte = page.extract_text()

                return texte if texte else "Aucun texte trouvé sur cette page."
        except Exception as e:
            return f"Erreur lors de la lecture du PDF : {e}"

    def extraire_chaine(self,texte):
        # Définir le pattern
        pattern = r"\d+\.\s[A-Z]+\s[A-Za-z]+\s[A-Z]+"
        # Rechercher la correspondance
        return re.findall(pattern, texte)

    def export_file(self):
        format_choisi=self.export_format.get()
        if format_choisi== "csv":
            # Demander à l'utilisateur de choisir le chemin et le nom du fichier de sortie
            export_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("Fichiers CSV", "*.csv")],
                title="Choisir le chemin et le nom du fichier de sortie"
            )

            if export_path:
                try:
                    # Ouvrir le fichier CSV en mode écriture
                    with open(export_path, "w", newline="", encoding="utf-8") as fichier_csv:
                        champs = ["Nom", "Sexe:MF", "Abreviation pays", "Rang", "Poids"]
                        writer = csv.DictWriter(fichier_csv, fieldnames=champs, delimiter=';')
                        writer.writeheader()

                        for res in self.resultats:
                            # Séparer la chaîne en parties
                            parts = res.split(" ")
                            longueur = len(parts)

                            # Extraire le rang, le nom, et le pays
                            rang = parts[0].replace(".", "")  # Enlever le point après le rang
                            nom = " ".join(parts[1:longueur - 1])  # Le nom se compose des éléments entre le rang et le pays
                            pays = parts[longueur - 1]  # Le pays est toujours à la fin

                            # Créer le dictionnaire pour écrire dans le CSV
                            donnees = {
                                "Nom": nom,
                                "Sexe:MF": "",
                                "Abreviation pays": pays,
                                "Rang": rang,
                                "Poids": ""  # Si vous avez un poids à ajouter, vous pouvez l'ajouter ici
                            }

                            # Ajouter les données dans le fichier CSV
                            writer.writerow(donnees)

                    self.txt_results.insert(tk.END, f"Les résultats ont été exportés dans le fichier {export_path}.\n")
                    messagebox.showinfo("Résultats exportés",f"Les résultats ont été exportés dans le fichier {export_path}.\n")
                except Exception as e:
                    messagebox.showerror("Erreur", f"Une erreur est survenue lors de l'exportation : {e}")
        elif format_choisi=="xlsx":
            # Demander à l'utilisateur de choisir le chemin et le nom du fichier de sortie
            export_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Fichiers Excel", "*.xlsx")],
                title="Choisir le chemin et le nom du fichier de sortie"
            )
            print("à faire")


    def clear_all(self):
        self.txt_results.delete('1.0', tk.END)
        self.btn_process.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)
        self.btn_clear.config(state=tk.DISABLED)

if __name__ == "__main__":
    app = Application()
    app.minsize(650, 300)
    app.iconbitmap("C:/extract.ico")
    app.mainloop()   #  pyinstaller --onefile --icon "extract.ico" --noconsole main.py
